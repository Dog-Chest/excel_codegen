"""渲染核心：组装上下文、渲染模板、导出代码文件。

Jinja2 环境与 ``pvs`` / ``wrap`` 过滤器在 :mod:`excel_codegen.jinja_env` 里
（``derived`` 求值也要用同一套语义），这里只是转出（re-export）。
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from jinja2 import Environment, TemplateError, TemplateNotFound, TemplateSyntaxError, UndefinedError, meta

from .excel_io import (
    check_required_sheets,
    check_value_constraints,
    load_workbook_file,
    read_cases,
    read_global_values,
    read_group_members,
)
from .jinja_env import build_environment, pvs, wrap
from .models import CaseData, ProjectConfig, RenderResult, TemplateDef
from .utils import ExcelError, RenderError, VarValue, column_index_to_letter, safe_filename, to_text

__all__ = [
    "FilterValue",
    "RenderOutput",
    "build_context",
    "build_environment",
    "case_matches",
    "check_asserts",
    "collect_variables",
    "compile_asserts",
    "compile_case_filter",
    "export_files",
    "filter_context",
    "member_of",
    "pvs",
    "render_all",
    "render_template",
    "template_environment",
    "validate_template",
    "wrap",
]


def _template_source(template: TemplateDef, base_dir: str | Path | None) -> str:
    """返回模板源码；``template_file`` 优先读取外部文件（相对**声明它的那个文件**所在目录）。

    ``extends`` 合并进来的模板会带自己的 ``source_dir`` —— 它的相对路径要相对原文件解析，
    而不是相对最终的项目 YAML。
    """
    if template.template_file:
        path = Path(template.template_file)
        base = template.source_dir or base_dir
        if base and not path.is_absolute():
            path = Path(base) / path
        if not path.exists():
            raise RenderError(f"模板 {template.name!r} 引用的模板文件不存在: {path}")
        try:
            return path.read_text(encoding="utf-8")
        except OSError as exc:
            raise RenderError(f"无法读取模板文件 {path}: {exc}") from exc
    return template.source_code


def template_environment(
    template: TemplateDef,
    base_dir: str | Path | None,
    env: Environment | None = None,
) -> Environment:
    """给某个模板准备 Jinja 环境。

    快照模式的模板可以 ``{% include "片段.j2" %}``，片段按**声明模板的那个文件**所在目录去找
    （``extends`` 进来的模板带自己的 ``source_dir``）。给了 ``env`` 就照用，不覆盖调用方的设置。
    """
    if env is not None:
        return env
    base = template.source_dir or base_dir
    return build_environment(search_path=[base] if base else None)


def _include_names(ast: Any) -> list[str]:
    """模板里 ``{% include %}`` 引用的片段名（动态名字取不到，跳过）。"""
    return [name for name in (meta.find_referenced_templates(ast) or []) if isinstance(name, str)]


def _check_includes(
    template: TemplateDef,
    ast: Any,
    environment: Environment,
    *,
    base_dir: str | Path | None,
    seen: frozenset[str] = frozenset(),
) -> list[tuple[str, Any]]:
    """检查 ``{% include %}`` 的片段都存在，并递归返回 ``[(片段名, 语法树), …]``。"""
    found: list[tuple[str, Any]] = []
    for name in _include_names(ast):
        if name in seen:
            continue
        if environment.loader is None:
            raise RenderError(
                f"模板 {template.name!r} 用了 {{% include {name!r} %}}，但当前 Jinja 环境没有搜索路径。"
                "（直接用库时请让 renderer 自己建环境：不要传 env，或者给 build_environment 传 search_path）"
            )
        try:
            source, _, _ = environment.loader.get_source(environment, name)
        except TemplateNotFound as exc:
            where = template.source_dir or base_dir or Path.cwd()
            raise RenderError(
                f"模板 {template.name!r} 的 {{% include {name!r} %}} 找不到片段："
                f"在 {where} 下没有 {name}（片段路径相对声明模板的那个文件，见指南 §17）"
            ) from exc
        try:
            child = environment.parse(source)
        except TemplateSyntaxError as exc:
            raise RenderError(
                f"片段 {name!r}（被模板 {template.name!r} include）语法错误：第 {exc.lineno} 行: {exc.message}"
            ) from exc
        found.append((name, child))
        found.extend(_check_includes(template, child, environment, base_dir=base_dir, seen=seen | {name}))
    return found


def validate_template(
    template: TemplateDef,
    *,
    env: Environment | None = None,
    base_dir: str | Path | None = None,
) -> None:
    """只做语法检查（解析模板、``{% include %}`` 片段、``case_filter``），不做渲染。"""
    environment = template_environment(template, base_dir, env)
    source = _template_source(template, base_dir)
    try:
        ast = environment.parse(source)
    except TemplateSyntaxError as exc:
        raise RenderError(f"模板 {template.name!r} 语法错误：第 {exc.lineno} 行: {exc.message}") from exc
    _check_includes(template, ast, environment, base_dir=base_dir)
    compile_case_filter(template, env=environment)


def collect_variables(
    template: TemplateDef,
    *,
    env: Environment | None = None,
    base_dir: str | Path | None = None,
) -> set[str]:
    """收集模板（含 ``{% include %}`` 片段与 ``case_filter``）中引用到的顶层变量名。"""
    environment = template_environment(template, base_dir, env)
    names: set[str] = set()
    try:
        ast = environment.parse(_template_source(template, base_dir))
        names |= set(meta.find_undeclared_variables(ast))
        # 片段里用到的变量也算"被引用" —— 否则 validate 会把它们误报成"定义了没人用"，
        # 也不会去检查它们有没有定义
        for _, fragment_ast in _check_includes(template, ast, environment, base_dir=base_dir):
            names |= set(meta.find_undeclared_variables(fragment_ast))
        if template.case_filter:
            # case_filter 是一段**表达式**，必须包进 {{ }} 才能按表达式解析
            filter_ast = environment.parse("{{ " + template.case_filter + " }}")
            names |= set(meta.find_undeclared_variables(filter_ast))
    except TemplateSyntaxError as exc:
        raise RenderError(f"模板 {template.name!r} 语法错误：第 {exc.lineno} 行: {exc.message}") from exc
    return names


# --------------------------------------------------------------------------- #
# case_filter：per-template 的 Case 过滤
# --------------------------------------------------------------------------- #
class FilterValue(str):
    """``case_filter`` 求值时的变量视图。

    它**像字符串**（内容等于组合值 ``prefix + value + suffix``，也就是 ``{{ x }}`` 的输出），
    同时又保留 ``.value`` / ``.prefix`` / ``.suffix``，所以下面两种写法都对::

        case_filter: "kind == 'EXT'"          # 组合值（无前后缀时就是值本身）
        case_filter: "draft.value > 20"       # 纯值，可做数值比较

    :class:`~excel_codegen.utils.VarValue` 本身是 dataclass，直接拿它和字符串比较永远不等，
    所以过滤器上下文用的是本类。
    """

    __slots__ = ("prefix", "suffix", "value")

    # 属性由 __new__ 填；显式标注是为了让类型检查看得见（__slots__ 本身不带类型信息）
    prefix: str
    suffix: str
    value: Any

    def __new__(cls, variable: VarValue) -> FilterValue:
        instance = super().__new__(cls, str(variable))
        instance.value = variable.value
        instance.prefix = variable.prefix
        instance.suffix = variable.suffix
        return instance

    def __repr__(self) -> str:  # pragma: no cover - 调试友好
        return f"FilterValue({str(self)!r}, value={self.value!r})"


def filter_context(context: Mapping[str, Any]) -> dict[str, Any]:
    """把渲染上下文里的 :class:`VarValue` 换成 :class:`FilterValue`（供 case_filter 使用）。"""
    return {key: FilterValue(value) if isinstance(value, VarValue) else value for key, value in context.items()}


def compile_case_filter(template: TemplateDef, *, env: Environment | None = None):
    """把 ``template.case_filter`` 编译成可调用表达式；未配置时返回 ``None``。"""
    if not template.case_filter:
        return None
    environment = env or build_environment()
    try:
        # undefined_to_none=False：变量缺失时报错，而不是静默判为 False
        return environment.compile_expression(template.case_filter, undefined_to_none=False)
    except TemplateSyntaxError as exc:
        raise RenderError(f"模板 {template.name!r} 的 case_filter 语法错误：第 {exc.lineno} 行: {exc.message}") from exc


def compile_asserts(config: ProjectConfig, *, env: Environment | None = None) -> list[tuple[str, Any]]:
    """编译根级 ``asserts``（跨变量校验），返回 ``[(原始表达式, 编译结果), …]``。

    语法错误在这里就报出来 —— 与 ``case_filter`` 同一套表达式机制，所以 ``x.value``
    那套写法与报错风格都一致。
    """
    if not config.asserts:
        return []
    environment = env or build_environment()
    compiled: list[tuple[str, Any]] = []
    for index, expression in enumerate(config.asserts, start=1):
        try:
            program = environment.compile_expression(expression, undefined_to_none=False)
        except TemplateSyntaxError as exc:
            raise RenderError(
                f"asserts 第 {index} 条语法错误：第 {exc.lineno} 行: {exc.message}（{expression!r}）"
            ) from exc
        compiled.append((expression, program))
    return compiled


def check_asserts(
    config: ProjectConfig,
    global_values: Mapping[str, VarValue],
    cases: Sequence[CaseData],
    *,
    members: Mapping[str, Mapping[str, VarValue]] | None = None,
    env: Environment | None = None,
) -> None:
    """对每个 Case 求值 ``asserts``；有假就抛 :class:`ExcelError`。

    单变量约束拦不住"吃水不能超过型深"这类**组合**错误，这一层补的就是它。
    """
    compiled = compile_asserts(config, env=env)
    if not compiled:
        return
    problems: list[str] = []
    for case in cases:
        # 成员值也要进上下文 —— 否则 asserts 里引用舱/设备的参数会报 undefined
        group_values = members[member_of(config, case, members)] if members else None
        context = build_context(global_values, case.values, case.name, group_values)
        for expression, program in compiled:
            try:
                passed = bool(program(**filter_context(context)))
            except TemplateError as exc:
                raise RenderError(
                    f"asserts {expression!r} 在 Case {case.name!r} 求值失败: {exc}"
                    "（表达式里引用的变量必须能在 Global / Local 表中取到）"
                ) from exc
            except Exception as exc:
                raise RenderError(
                    f"asserts {expression!r} 在 Case {case.name!r} 求值失败: {exc}"
                    "（数值比较请写 x.value，例如 draft.value <= d_tank.value）"
                ) from exc
            if not passed:
                column = column_index_to_letter(case.column)
                problems.append(f"  Case '{case.name}'（Local 第 {column} 列）不满足：{expression}")
    if problems:
        raise ExcelError(
            f"参数不满足 YAML 里的 asserts（跨变量校验），共 {len(problems)} 处：\n"
            + "\n".join(problems)
            + "\n  → 改 Excel 里的取值，或调整 YAML 里的 asserts"
        )


def case_matches(
    expression,
    context: Mapping[str, Any],
    *,
    template_name: str,
) -> bool:
    """对单个 Case 求值 ``case_filter``。表达式为空/未配置时一律返回 ``True``。"""
    if expression is None:
        return True
    try:
        return bool(expression(**filter_context(context)))
    except TemplateError as exc:
        raise RenderError(
            f"模板 {template_name!r} 的 case_filter 求值失败: {exc}"
            "（表达式里引用的变量必须能在 Global / Local 表中取到）"
        ) from exc
    except Exception as exc:
        raise RenderError(
            f"模板 {template_name!r} 的 case_filter 求值失败: {exc}（数值比较请写 x.value，例如 draft.value > 20）"
        ) from exc


# --------------------------------------------------------------------------- #
# 上下文与渲染
# --------------------------------------------------------------------------- #
def build_context(
    global_values: Mapping[str, VarValue],
    local_values: Mapping[str, VarValue],
    case_name: str,
    group_values: Mapping[str, VarValue] | None = None,
) -> dict[str, Any]:
    """组装单个 Case 的渲染上下文。

    优先级从低到高：**成员（group）→ 全局（global）→ 局部（local）→ case_name**。
    成员值放最低是因为它描述的是"这个舱/设备是什么"，而 global / local 是"这次计算怎么算"。
    """
    context: dict[str, Any] = {}
    if group_values:
        context.update(group_values)
    context.update(global_values)
    context.update(local_values)
    context["case_name"] = case_name
    context["template_name"] = ""  # 渲染具体模板时会被覆盖
    return context


def member_of(
    config: ProjectConfig,
    case: CaseData,
    members: Mapping[str, Mapping[str, VarValue]],
) -> str:
    """从 Case 的 ``key`` 变量取出成员名，并确认成员表里真有这个成员。"""
    group = config.group
    assert group is not None  # 调用方负责只在有成员表时调用
    value = case.values.get(group.key)
    name = to_text(value.text).strip() if value is not None else ""
    if not name:
        raise ExcelError(f"Case {case.name!r} 的 {group.key!r} 是空的 —— 它要指向成员表 {group.sheet!r} 里的一个成员")
    if name not in members:
        available = ", ".join(members) or "（成员表是空的）"
        raise ExcelError(
            f"Case {case.name!r} 的 {group.key!r} = {name!r}，但成员表 {group.sheet!r} 里没有这个成员"
            f"（可选: {available}）"
        )
    return name


def render_template(
    template: TemplateDef,
    context: Mapping[str, Any],
    *,
    env: Environment | None = None,
    base_dir: str | Path | None = None,
) -> str:
    """渲染单个模板，并把 Jinja2 异常翻译成清晰的中文提示。"""
    environment = template_environment(template, base_dir, env)
    source = _template_source(template, base_dir)
    try:
        compiled = environment.from_string(source)
    except TemplateSyntaxError as exc:
        raise RenderError(f"模板 {template.name!r} 语法错误：第 {exc.lineno} 行: {exc.message}") from exc

    data = dict(context)
    data["template_name"] = template.name
    try:
        return compiled.render(**data)
    except UndefinedError as exc:
        raise RenderError(
            f"模板 {template.name!r} 变量缺失: {exc.message or exc}"
            "（请检查 YAML 的 variables 与 Excel 中的变量名是否一致）"
        ) from exc
    except TemplateError as exc:
        raise RenderError(f"模板 {template.name!r} 渲染失败: {exc}") from exc


@dataclass
class RenderOutput:
    """一次完整渲染的结果集合。"""

    global_values: dict[str, VarValue]
    cases: list[CaseData]
    results: dict[str, list[RenderResult]] = field(default_factory=dict)
    #: 被 ``case_filter`` 跳过的 Case：``{模板名: [Case 名, ...]}``
    skipped: dict[str, list[str]] = field(default_factory=dict)
    #: 值得提醒但不致命的问题（CLI 会打出来）
    warnings: list[str] = field(default_factory=list)

    def iter_results(self) -> Iterator[RenderResult]:
        for per_case in self.results.values():
            yield from per_case

    def total(self) -> int:
        return sum(len(item) for item in self.results.values())

    def skipped_total(self) -> int:
        return sum(len(item) for item in self.skipped.values())

    def preview(self, *, max_cases: int | None = None) -> str:
        """把结果拼成一段纯文本预览（CLI / 调试用）。"""
        blocks: list[str] = []
        for template_name, per_case in self.results.items():
            cases = per_case if max_cases is None else per_case[:max_cases]
            for result in cases:
                blocks.append(f"### {template_name} · {result.case_name}\n{result.text.rstrip()}")
        return "\n\n".join(blocks)


def render_all(
    config: ProjectConfig,
    excel_path: str | Path,
    *,
    env: Environment | None = None,
    only_cases: Sequence[str] | None = None,
) -> RenderOutput:
    """读取 Excel 中填写的参数，按 Case 渲染所有模板。

    ``template.case_filter`` 可以只让某个模板作用于部分 Case（例如一本工作簿里
    同时放内外压两套模板，用 ``kind`` 区分）。
    """
    workbook = load_workbook_file(excel_path)
    read_warnings: list[str] = []
    try:
        check_required_sheets(workbook, config)
        global_values = read_global_values(workbook, config, warnings=read_warnings)
        cases = read_cases(workbook, config, global_values=global_values, warnings=read_warnings)
        members = read_group_members(workbook, config, warnings=read_warnings)
    finally:
        workbook.close()

    # 取值约束（min / max / choices / pattern）：声明了就一定查，别让手滑的数字生成出错误代码
    check_value_constraints(config, global_values, cases, members=members)
    # 跨变量校验（asserts）：拦"吃水不能超过型深"这类组合错误
    check_asserts(config, global_values, cases, members=members)

    if only_cases:
        wanted = [to_text(name).strip() for name in only_cases]
        by_name = {case.name: case for case in cases}
        missing = [name for name in wanted if name not in by_name]
        if missing:
            raise ExcelError(f"Excel 中不存在这些 Case: {', '.join(missing)}（可用: {', '.join(by_name)}）")
        cases = [by_name[name] for name in wanted]

    results: dict[str, list[RenderResult]] = {}
    skipped: dict[str, list[str]] = {}
    environment_cache: dict[str, Environment] = {}
    for template in config.templates:
        # 每个模板用"声明它的那个目录"的环境：{% include %} 的片段相对那里解析
        base = template.source_dir or config.source_dir
        key = str(base or "")
        if key not in environment_cache:
            environment_cache[key] = template_environment(template, config.source_dir, env)
        environment = environment_cache[key]
        expression = compile_case_filter(template, env=environment)
        per_case: list[RenderResult] = []
        skipped_cases: list[str] = []
        for case in cases:
            group_values = members[member_of(config, case, members)] if config.group is not None else None
            context = build_context(global_values, case.values, case.name, group_values)
            if not case_matches(expression, context, template_name=template.name):
                skipped_cases.append(case.name)
                continue
            text = render_template(template, context, env=environment, base_dir=config.source_dir)
            per_case.append(RenderResult(template.name, case.name, text))
        results[template.name] = per_case
        if skipped_cases:
            skipped[template.name] = skipped_cases

    return RenderOutput(
        global_values=dict(global_values),
        cases=list(cases),
        results=results,
        skipped=skipped,
        warnings=[*_case_warnings(config, cases), *read_warnings],
    )


def _case_warnings(config: ProjectConfig, cases: Sequence[CaseData]) -> list[str]:
    """全空的 Case 列值得提醒：所有变量回落到 YAML ``default``，
    而 ``case_filter`` 会按 default 把它划进某个规则集 —— 表上却只是一个空列。"""
    warnings: list[str] = []
    blank = [case.name for case in cases if not case.explicit_values]
    if not blank:
        return warnings
    filtered = [t.name for t in config.templates if t.case_filter]
    warnings.append(
        f"这些 Case 列在 Local 表里是空的：{', '.join(blank)} —— "
        "所有局部变量都取自 YAML default"
        + (
            f"；`case_filter` 会按 default 把它们划进某个规则集（{', '.join(filtered)}），"
            "新插的空列请先填上归属变量（例如 kind）"
            if filtered
            else ""
        )
    )
    return warnings


# --------------------------------------------------------------------------- #
# 导出代码文件
# --------------------------------------------------------------------------- #
def export_files(
    config: ProjectConfig,
    results: Mapping[str, Sequence[RenderResult]],
    outdir: str | Path,
    *,
    env: Environment | None = None,
    overwrite: bool = True,
) -> list[Path]:
    """把渲染结果导出为代码文件。

    文件名由 ``template.filename`` 决定（支持 Jinja2，可用 ``{{ case_name }}`` /
    ``{{ template_name }}``）；未配置时默认 ``<模板名>_<Case名><extension>``。
    """
    environment = env or build_environment()
    directory = Path(outdir)
    directory.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    for template in config.templates:
        per_case = results.get(template.name)
        if not per_case:
            continue
        pattern = template.filename or (template.name + "_{{ case_name }}" + template.extension)
        for result in per_case:
            try:
                raw_name = environment.from_string(pattern).render(
                    case_name=result.case_name,
                    template_name=template.name,
                )
            except TemplateError as exc:
                raise RenderError(f"模板 {template.name!r} 的 filename 渲染失败: {exc}") from exc
            file_name = safe_filename(raw_name)
            if not file_name:
                raise RenderError(f"模板 {template.name!r} 在 Case {result.case_name!r} 下生成了空文件名")
            target = directory / file_name
            if target.exists() and not overwrite:
                raise RenderError(f"目标文件已存在: {target}（需要覆盖请去掉 --no-overwrite）")
            target.parent.mkdir(parents=True, exist_ok=True)
            text = result.text if result.text.endswith("\n") else result.text + "\n"
            try:
                target.write_text(text, encoding="utf-8")
            except OSError as exc:
                raise RenderError(f"无法写入文件 {target}: {exc}") from exc
            written.append(target)
    return written
