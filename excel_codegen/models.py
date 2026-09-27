"""Pydantic 数据模型：YAML 配置结构 + 运行期数据（CaseData / RenderResult）。

配置文件结构::

    version: 1
    excel:
      output: "template.xlsx"
      template_sheet: "Template"
      sheets:
        global: "Global Parameter"
        local: "Local Parameter"
        outputs: ["Output"]
    variables:
      global: [...]
      local: [...]
    templates: [...]
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    ValidationError,
    field_validator,
    model_validator,
)

from .utils import (
    ConfigError,
    VarValue,
    is_identifier,
    parse_cell,
    split_lines,
    to_text,
)

__all__ = [
    "FIRST_CASE_COLUMN",
    "CaseData",
    "Direction",
    "ExcelConfig",
    "ExcelSheets",
    "ProjectConfig",
    "RenderResult",
    "TemplateDef",
    "VarType",
    "VariableDef",
    "VariablesConfig",
    "format_validation_error",
    "load_config",
]

#: Local Parameter 工作表中第一个 Case 列（E 列），A-D 为变量名/描述/Prefix/Suffix。
FIRST_CASE_COLUMN: int = 5

Direction = Literal["horizontal", "vertical"]
VarType = Literal["auto", "string", "int", "float", "bool", "raw"]
#: ``snapshot`` = 脚本渲染后把**文本**写进输出表（默认）；
#: ``excel``    = 把模板编译成**Excel 公式**写进输出表，改参数由 Excel 自己重算。
Engine = Literal["snapshot", "excel"]

#: 渲染上下文里由程序注入、不允许作为变量名使用的保留名。
RESERVED_NAMES: frozenset[str] = frozenset({"case_name", "template_name"})


def _duplicates(items: list[str]) -> list[str]:
    seen: set[str] = set()
    dupes: list[str] = []
    for item in items:
        if item in seen and item not in dupes:
            dupes.append(item)
        seen.add(item)
    return dupes


def _valid_sheet_name(value: Any) -> str:
    """校验并规范化一个 Excel 工作表名。"""
    name = to_text(value).strip()
    if not name:
        raise ValueError("工作表名不能为空")
    if len(name) > 31 or any(char in name for char in "[]:*?/\\"):
        raise ValueError(f"工作表名 {name!r} 非法（Excel 限制：<=31 字符且不能含 []:*?/\\）")
    return name


def format_validation_error(exc: ValidationError) -> str:
    """把 pydantic 的校验错误整理成多行、便于阅读的中文提示。"""
    lines: list[str] = []
    for error in exc.errors():
        location = ".".join(str(part) for part in error.get("loc", ())) or "<root>"
        lines.append(f"  - {location}: {error.get('msg', '非法取值')}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 变量定义
# --------------------------------------------------------------------------- #
class VariableDef(BaseModel):
    """一个变量的定义（名称、描述、默认值、前缀、后缀、类型）。"""

    model_config = ConfigDict(extra="forbid")

    name: str
    description: str = ""
    default: Any = ""
    prefix: str = ""
    suffix: str = ""
    type: VarType = "auto"
    #: 派生参数：一段 Jinja2 **表达式**，只能引用 ``global`` 与**同一个 Case** 的 ``local``。
    #: 例如 ``derived: "k_c * h_di"``。派生参数不用在 Excel 里填值（那一格由工具写成公式或算好的值）。
    derived: str | None = None

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("变量名不能为空")
        if not is_identifier(name):
            raise ValueError(
                f"变量名 {name!r} 不是合法标识符（需匹配 [A-Za-z_][A-Za-z0-9_]*），因为要在 Jinja2 模板中直接引用"
            )
        if name in RESERVED_NAMES:
            raise ValueError(f"变量名 {name!r} 是保留字（渲染时由程序注入），请改名")
        return name

    @field_validator("description", "prefix", "suffix", mode="before")
    @classmethod
    def _to_str(cls, value: Any) -> str:
        return "" if value is None else to_text(value)

    @field_validator("derived", mode="before")
    @classmethod
    def _derived_to_str(cls, value: Any) -> str | None:
        if value is None:
            return None
        text = to_text(value).strip()
        return text or None

    @model_validator(mode="after")
    def _check_derived(self) -> "VariableDef":
        if self.derived and to_text(self.default) != "":
            raise ValueError(
                f"变量 {self.name!r} 同时写了 derived 与 default —— 派生参数的值由表达式算出来，"
                "不能同时给它一个填写值；请删掉其中一个"
            )
        return self

    @property
    def is_derived(self) -> bool:
        return bool(self.derived)


class VariablesConfig(BaseModel):
    """全局变量 + 局部变量定义。"""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    global_: list[VariableDef] = Field(default_factory=list, alias="global")
    local: list[VariableDef] = Field(default_factory=list)

    @property
    def global_names(self) -> list[str]:
        return [item.name for item in self.global_]

    @property
    def local_names(self) -> list[str]:
        return [item.name for item in self.local]


# --------------------------------------------------------------------------- #
# Excel 结构
# --------------------------------------------------------------------------- #
class ExcelSheets(BaseModel):
    """工作表名称映射。"""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    global_: str = Field("Global Parameter", alias="global")
    local: str = "Local Parameter"
    outputs: list[str] = Field(default_factory=lambda: ["Output"])

    @field_validator("global_", "local")
    @classmethod
    def _check_sheet_name(cls, value: str) -> str:
        return _valid_sheet_name(value)

    @field_validator("outputs")
    @classmethod
    def _check_outputs(cls, value: list[str]) -> list[str]:
        names = [to_text(item).strip() for item in value]
        if not names or any(not item for item in names):
            raise ValueError("excel.sheets.outputs 至少需要一个非空的工作表名")
        dupes = _duplicates(names)
        if dupes:
            raise ValueError(f"excel.sheets.outputs 中工作表名重复: {', '.join(dupes)}")
        return names


class ExcelConfig(BaseModel):
    """Excel 模板文件的整体配置。"""

    model_config = ConfigDict(extra="forbid")

    output: str = "template.xlsx"
    template_sheet: str | None = "Template"
    #: 工作簿里的"使用说明"表（放在第一张）。写清三步、命令、生成时间与指纹；
    #: 由 create_template 生成、write_results 刷新。设为 null 则不生成。
    howto_sheet: str | None = "HOWTO"
    sheets: ExcelSheets = Field(default_factory=ExcelSheets)

    @field_validator("output")
    @classmethod
    def _check_output(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("excel.output 不能为空")
        return name

    @field_validator("template_sheet", "howto_sheet")
    @classmethod
    def _check_optional_sheet_name(cls, value: str | None) -> str | None:
        if value is None or to_text(value).strip() == "":
            return None
        return _valid_sheet_name(value)


# --------------------------------------------------------------------------- #
# 模板定义
# --------------------------------------------------------------------------- #
class TemplateDef(BaseModel):
    """一个输出模板：源码 + 输出位置 + 布局方向。"""

    model_config = ConfigDict(extra="forbid")

    name: str
    description: str = ""
    output_sheet: str = "Output"
    start_cell: str = "B2"
    direction: Direction = "horizontal"
    write_case_headers: bool = True
    filename: str | None = None
    extension: str = ".txt"
    code: str | None = None
    template_file: str | None = None
    #: 输出引擎：``snapshot``（默认，写文本快照）或 ``excel``（写公式，改参数不用重跑脚本）。
    #: ``excel`` 只支持"纯替换"子集：{{ x }} / {{ x.value }} / {{ x.prefix }} / {{ x.suffix }} /
    #: {{ case_name }} / {{ template_name }}；含控制流或过滤器会报错。
    engine: Engine = "snapshot"
    #: 可选：Jinja2 表达式，对每个 Case 的上下文求值；为假则该模板跳过这个 Case。
    #: 例如 ``case_filter: "kind == 'EXT'"``。用于"一本工作簿放两套规则"的场景。
    case_filter: str | None = None

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("模板名不能为空")
        return name

    @field_validator("description", mode="before")
    @classmethod
    def _description_to_str(cls, value: Any) -> str:
        return "" if value is None else to_text(value)

    @field_validator("case_filter", mode="before")
    @classmethod
    def _case_filter_to_str(cls, value: Any) -> str | None:
        if value is None:
            return None
        text = to_text(value).strip()
        return text or None

    @field_validator("extension")
    @classmethod
    def _check_extension(cls, value: str) -> str:
        ext = to_text(value).strip()
        if not ext:
            return ""
        return ext if ext.startswith(".") else f".{ext}"

    @field_validator("start_cell")
    @classmethod
    def _check_start_cell(cls, value: str) -> str:
        try:
            parse_cell(value)
        except ValueError as exc:  # pragma: no cover - 信息透传
            raise ValueError(str(exc)) from exc
        return to_text(value).strip().upper()

    @model_validator(mode="after")
    def _check_source(self) -> "TemplateDef":
        if not (self.code and self.code.strip()) and not self.template_file:
            raise ValueError("必须提供 code（内联模板）或 template_file（外部模板文件）之一")
        if self.filename and not self.filename.strip():
            raise ValueError("filename 不能为空字符串（如需默认命名请删除该字段）")
        return self

    @property
    def source_code(self) -> str:
        """内联模板源码（``template_file`` 场景下为空字符串）。"""
        return self.code or ""


# --------------------------------------------------------------------------- #
# 顶层配置
# --------------------------------------------------------------------------- #
class ProjectConfig(BaseModel):
    """整个项目的配置根对象。"""

    model_config = ConfigDict(extra="forbid")

    version: int = 1
    excel: ExcelConfig = Field(default_factory=ExcelConfig)
    variables: VariablesConfig = Field(default_factory=VariablesConfig)
    templates: list[TemplateDef]

    #: 配置文件所在目录（加载时自动填充，用于解析相对路径的 template_file）。
    source_dir: Path | None = Field(default=None, exclude=True)

    @model_validator(mode="after")
    def _validate_config(self) -> "ProjectConfig":
        if self.version < 1:
            raise ValueError("version 必须 >= 1")
        if not self.templates:
            raise ValueError("至少需要定义一个 template")

        template_names = [item.name for item in self.templates]
        dupes = _duplicates(template_names)
        if dupes:
            raise ValueError(f"template 名称重复: {', '.join(dupes)}")

        for scope, names in (
            ("global", self.variables.global_names),
            ("local", self.variables.local_names),
        ):
            dupes = _duplicates(names)
            if dupes:
                raise ValueError(f"{scope} 变量名重复: {', '.join(dupes)}")

        outputs = set(self.excel.sheets.outputs)
        for template in self.templates:
            if template.output_sheet not in outputs:
                raise ValueError(
                    f"template {template.name!r} 的 output_sheet {template.output_sheet!r} "
                    f"未在 excel.sheets.outputs 中声明（当前: {', '.join(self.excel.sheets.outputs)}）"
                )

        used = {self.excel.sheets.global_, self.excel.sheets.local, *outputs}
        for label, sheet_name in (
            ("excel.template_sheet", self.excel.template_sheet),
            ("excel.howto_sheet", self.excel.howto_sheet),
        ):
            if sheet_name and sheet_name in used:
                raise ValueError(
                    f"{label} {sheet_name!r} 与 Global / Local / Output 表名冲突，请改名或设为 null"
                )
        if (
            self.excel.template_sheet
            and self.excel.howto_sheet
            and self.excel.template_sheet == self.excel.howto_sheet
        ):
            raise ValueError(
                f"excel.template_sheet 与 excel.howto_sheet 不能同名（都是 {self.excel.template_sheet!r}）"
            )
        return self

    # -- 便捷访问 ---------------------------------------------------------- #
    @property
    def global_variables(self) -> list[VariableDef]:
        return list(self.variables.global_)

    @property
    def local_variables(self) -> list[VariableDef]:
        return list(self.variables.local)

    @property
    def global_names(self) -> list[str]:
        return self.variables.global_names

    @property
    def local_names(self) -> list[str]:
        return self.variables.local_names

    @property
    def defined_names(self) -> set[str]:
        """所有已定义变量名（用于校验模板引用的变量是否存在）。"""
        return set(self.global_names) | set(self.local_names) | set(RESERVED_NAMES)

    def template_by_name(self, name: str) -> TemplateDef | None:
        for template in self.templates:
            if template.name == name:
                return template
        return None


class _StrictLoader(yaml.SafeLoader):
    """SafeLoader + 同一映射里的重复键检测。

    PyYAML 默认对重复键**不报错**（后者覆盖前者），于是 YAML 里写两个 ``variables:``
    会让第一个整块静默消失，错误要到渲染期才以"变量缺失"的形式暴露。
    """


def _construct_mapping_strict(loader: yaml.SafeLoader, node: yaml.MappingNode, deep: bool = False) -> dict:
    mapping: dict = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        try:
            duplicated = key in mapping
        except TypeError as exc:  # 不可哈希的复杂键
            raise ConfigError(
                f"YAML 键非法（不可哈希）: 第 {key_node.start_mark.line + 1} 行"
            ) from exc
        if duplicated:
            raise ConfigError(
                f"YAML 键重复: {key!r}（第 {key_node.start_mark.line + 1} 行）"
                "—— 同一个映射里同一个键只能出现一次（重复会让前一份被静默覆盖）"
            )
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_StrictLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _construct_mapping_strict
)


def load_config(path: str | Path) -> ProjectConfig:
    """读取并校验 YAML 配置，失败时抛出带清晰提示的 :class:`ConfigError`。"""
    config_path = Path(path)
    if not config_path.exists():
        raise ConfigError(f"配置文件不存在: {config_path}")
    if config_path.is_dir():
        raise ConfigError(f"配置路径是目录而不是文件: {config_path}")

    try:
        raw_text = config_path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigError(f"无法读取配置文件 {config_path}: {exc}") from exc

    try:
        raw = yaml.load(raw_text, Loader=_StrictLoader)
    except ConfigError:
        raise
    except yaml.YAMLError as exc:
        mark = getattr(exc, "problem_mark", None)
        where = f"（第 {mark.line + 1} 行，第 {mark.column + 1} 列）" if mark else ""
        raise ConfigError(f"YAML 解析失败 {config_path}{where}:\n  {exc}") from exc

    if raw is None:
        raise ConfigError(f"配置文件内容为空: {config_path}")
    if not isinstance(raw, dict):
        raise ConfigError(f"配置文件根节点必须是映射（mapping），实际是 {type(raw).__name__}: {config_path}")

    try:
        config = ProjectConfig.model_validate(raw)
    except ValidationError as exc:
        raise ConfigError(
            f"配置校验失败 {config_path}:\n{format_validation_error(exc)}"
        ) from exc

    config.source_dir = config_path.resolve().parent
    return config


# --------------------------------------------------------------------------- #
# 运行期数据
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class CaseData:
    """Local Parameter 表中一列 Case 的取值。"""

    name: str
    column: int
    values: dict[str, VarValue]
    #: 这一列在表里是否**至少填过一个**局部变量的值；
    #: 全空时所有变量都会回落 YAML ``default``（新插的空列就此"悄悄"落进某个规则集）。
    explicit_values: bool = True

    @property
    def index(self) -> int:
        """Case 序号（Case1 -> 1）。"""
        return self.column - FIRST_CASE_COLUMN + 1


@dataclass(frozen=True)
class RenderResult:
    """某个模板在某个 Case 下的渲染结果。"""

    template_name: str
    case_name: str
    text: str

    @property
    def lines(self) -> list[str]:
        """渲染结果的行列表（不含结尾空行），用于写入 Excel。"""
        return split_lines(self.text)

    @property
    def line_count(self) -> int:
        return len(self.lines)
