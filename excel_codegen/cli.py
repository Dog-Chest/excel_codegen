"""命令行接口：``excel-codegen init | render | validate | check``。

控制台注意事项：中文 Windows 的 GBK 控制台编不出 ``✓`` / ``✗`` 之类的字符，
所以本模块只用 ASCII 标记（``OK`` / ``ERROR`` / ``!``），并在导入时给
stdout/stderr 加上 ``errors="backslashreplace"`` 兜底 —— 成功路径绝不允许
因为"最后一行字打不出来"而返回非 0 退出码。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional, Sequence

import typer
from rich.console import Console
from rich.panel import Panel
from rich.syntax import Syntax
from rich.table import Table

from . import __version__
from .excel_io import (
    case_column_map,
    check_required_sheets,
    check_value_constraints,
    create_template,
    input_fingerprint,
    load_workbook_file,
    output_fingerprint,
    read_cases,
    read_global_values,
    read_metadata,
    template_source,
    write_results,
)
from .derived import expression_names
from .derived import validate_config as derived_validate_config
from .formula import LONG_FORMULA_WARN, compile_formulas
from .formula_eval import FormulaEvalError, evaluate_template_values
from .models import FIRST_CASE_COLUMN, ProjectConfig, RenderResult, load_config
from .renderer import (
    RenderOutput,
    build_environment,
    collect_variables,
    export_files,
    render_all,
    validate_template,
)
from .utils import CodeGenError, ExcelError, column_index_to_letter, parse_cell, to_text


def _make_streams_forgiving() -> None:
    """让 stdout/stderr 遇到当前编码表达不了的字符时降级，而不是抛异常。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(errors="backslashreplace")
        except (ValueError, OSError):  # pragma: no cover - 取决于宿主环境
            pass


_make_streams_forgiving()

app = typer.Typer(
    add_completion=False,
    no_args_is_help=True,
    help="Excel 模板参数填写 + Jinja2 代码生成器（支持 Prefix/Value/Suffix 组合）",
)
console = Console()
error_console = Console(stderr=True)


# --------------------------------------------------------------------------- #
# 公共小工具
# --------------------------------------------------------------------------- #
def _fail(message: object, code: int = 1) -> typer.Exit:
    error_console.print(f"[bold red]ERROR[/] {message}")
    return typer.Exit(code=code)


def _warn(message: str) -> None:
    console.print(f"[bold yellow]![/] {message}")


def _version_callback(value: bool) -> None:
    if value:
        console.print(f"excel-codegen [cyan]{__version__}[/]")
        raise typer.Exit()


def _resolve_excel_path(config: ProjectConfig, excel: Optional[Path]) -> Path:
    """--excel 优先；否则使用配置里的 excel.output（相对当前工作目录）。"""
    if excel is not None:
        return Path(excel)
    return Path(config.excel.output)


def _parse_cases(value: str) -> int | list[str]:
    """``--cases`` 既接受数量（``3``）也接受名字列表（``EXT-T20,INT-T15``）。"""
    text = to_text(value).strip()
    if not text:
        raise ValueError("--cases 不能为空")
    if text.lstrip("+-").isdigit():
        count = int(text)
        if count < 1:
            raise ValueError(f"--cases 的数量至少为 1（收到 {count}）")
        return count
    names = [item.strip() for item in text.split(",")]
    if any(not name for name in names):
        raise ValueError("--cases 里的 Case 名不能为空，例如 --cases Case1,Case2")
    return names


def _case_names_of(spec: int | list[str]) -> list[str]:
    return [f"Case{index}" for index in range(1, spec + 1)] if isinstance(spec, int) else list(spec)


def _open_excel(project: ProjectConfig, excel: Optional[Path]) -> Path:
    path = _resolve_excel_path(project, excel)
    if not path.exists():
        raise ExcelError(
            f"Excel 文件不存在: {path}（请先运行 `excel-codegen init` 生成模板并填写参数）"
        )
    return path


@app.callback()
def main(
    version: bool = typer.Option(
        False,
        "--version",
        callback=_version_callback,
        is_eager=True,
        help="显示版本号并退出",
    ),
) -> None:
    """Excel 模板参数填写 + Jinja2 代码生成器。"""


# --------------------------------------------------------------------------- #
# init
# --------------------------------------------------------------------------- #
@app.command("init")
def init_command(
    config: Path = typer.Option(
        ...,
        "--config",
        "-c",
        exists=True,
        dir_okay=False,
        readable=True,
        help="YAML 配置文件路径",
    ),
    output: Optional[Path] = typer.Option(
        None, "--output", "-o", help="输出的 Excel 模板路径（默认取配置中的 excel.output）"
    ),
    cases: str = typer.Option(
        "2",
        "--cases",
        help="初始 Case 列：数量（如 3）或逗号分隔的名字（如 EXT-T20,INT-T15）",
    ),
    force: bool = typer.Option(False, "--force", "-f", help="目标 Excel 已存在时覆盖"),
    template_sheet: bool = typer.Option(
        True,
        "--template-sheet/--no-template-sheet",
        help="是否生成隐藏的 Template 表（保存模板原文）",
    ),
    howto_sheet: bool = typer.Option(
        True,
        "--howto/--no-howto",
        help="是否生成 HOWTO 表（放在第一张，写清下一步跑什么）",
    ),
) -> None:
    """根据 YAML 生成 Excel 参数填写模板。"""
    try:
        spec = _parse_cases(cases)
        project = load_config(config)
        target = create_template(
            project,
            output or Path(project.excel.output),
            cases=spec,
            overwrite=force,
            include_template_sheet=template_sheet,
            include_howto_sheet=howto_sheet,
        )
    except CodeGenError as exc:
        raise _fail(exc) from exc
    except (ValueError, OSError) as exc:
        raise _fail(exc) from exc

    table = Table(title="Excel 模板已生成", show_header=True, header_style="bold cyan")
    table.add_column("项目", style="bold")
    table.add_column("内容")
    table.add_row("文件", str(target))
    table.add_row("Global 表", project.excel.sheets.global_)
    table.add_row("Local 表", f"{project.excel.sheets.local}（Case 列：E 起）")
    table.add_row("Output 表", ", ".join(project.excel.sheets.outputs))
    table.add_row("Case 列", ", ".join(_case_names_of(spec)))
    table.add_row("说明表", project.excel.howto_sheet or "（未生成）")
    table.add_row("模板数", str(len(project.templates)))
    console.print(table)
    console.print(
        Panel(
            "[bold]下一步[/]\n"
            f"1. 在 [cyan]{project.excel.sheets.global_}[/] 表填写 B 列（Value），D/E 列可覆盖 Prefix/Suffix；\n"
            f"2. 在 [cyan]{project.excel.sheets.local}[/] 表从 E 列开始按 Case 填写，如需更多 Case 直接右拉复制；\n"
            "3. 运行渲染（不加 --write-excel 只预览，不会改动 Excel）：\n"
            f"   [green]excel-codegen render --config {config} --excel {target} --write-excel[/]\n"
            f"   [green]excel-codegen render --config {config} --excel {target} --outdir generated/[/]",
            title="使用说明",
            border_style="green",
        )
    )


# --------------------------------------------------------------------------- #
# render
# --------------------------------------------------------------------------- #
@app.command("render")
def render_command(
    config: Path = typer.Option(
        ...,
        "--config",
        "-c",
        exists=True,
        dir_okay=False,
        readable=True,
        help="YAML 配置文件路径",
    ),
    excel: Optional[Path] = typer.Option(
        None, "--excel", "-x", help="已填写的 Excel 模板（默认取配置中的 excel.output）"
    ),
    write_excel: bool = typer.Option(
        False, "--write-excel", "-w", help="把渲染结果写入 Excel 的 Output 表"
    ),
    outdir: Optional[Path] = typer.Option(
        None, "--outdir", "-d", help="把渲染结果导出为代码文件到该目录"
    ),
    case: Optional[list[str]] = typer.Option(
        None, "--case", help="只渲染指定 Case（可重复传入）"
    ),
    show: bool = typer.Option(True, "--show/--no-show", help="在终端打印渲染结果"),
    overwrite: bool = typer.Option(
        True, "--overwrite/--no-overwrite", help="导出文件已存在时是否覆盖"
    ),
) -> None:
    """读取填好的 Excel + YAML，渲染模板并输出到 Excel / 代码文件 / 终端。"""
    try:
        project = load_config(config)
        excel_path = _open_excel(project, excel)
        output = render_all(project, excel_path, only_cases=list(case) if case else None)

        written_files: list[Path] = []
        output_warnings: list[str] = []
        command_text = (
            f"excel-codegen render -c {config} -x {excel_path}"
            + (" --write-excel" if write_excel else "")
            + (f" --outdir {outdir}" if outdir is not None else "")
        )
        if write_excel:
            write_results(
                excel_path,
                project,
                output.results,
                command=command_text,
                warnings=output_warnings,
            )
        if outdir is not None:
            written_files = export_files(
                project, output.results, outdir, overwrite=overwrite
            )
    except CodeGenError as exc:
        raise _fail(exc) from exc
    except (ValueError, OSError) as exc:
        raise _fail(exc) from exc

    if show:
        for template in project.templates:
            for result in output.results.get(template.name, []):
                console.print(
                    Panel(
                        Syntax(result.text.rstrip("\n") or " ", "text", theme="ansi_dark", word_wrap=True),
                        title=f"[bold]{template.name}[/] · [cyan]{result.case_name}[/]",
                        subtitle=f"→ {template.output_sheet} @ {template.start_cell} ({template.direction})",
                        border_style="cyan",
                    )
                )

    summary = Table(title="渲染摘要", header_style="bold cyan")
    summary.add_column("项目", style="bold")
    summary.add_column("内容")
    summary.add_row("配置文件", str(config))
    summary.add_row("Excel", str(excel_path))
    summary.add_row("Case", ", ".join(item.name for item in output.cases))
    summary.add_row("渲染结果", f"{output.total()} 个（{len(project.templates)} 模板 × {len(output.cases)} Case）")
    if output.skipped:
        detail = "；".join(f"{name} 跳过 {', '.join(cases)}" for name, cases in output.skipped.items())
        summary.add_row("case_filter", f"跳过 {output.skipped_total()} 个（{detail}）")
    # 永远显示这一行：让"成功"与"成功但没动文件"能一眼分开
    summary.add_row("写回 Excel", "是" if write_excel else "否（需要 --write-excel）")
    summary.add_row("导出文件", "\n".join(str(item) for item in written_files) if written_files else "无")
    console.print(summary)

    if not write_excel and outdir is None:
        _warn("本次只预览：没有写回 Excel，也没有导出文件。加 --write-excel / --outdir 才会落盘。")
    elif not write_excel:
        _warn(f"没有写回 Excel（只导出了文件）：Excel 里的 {project.excel.sheets.outputs[0]} 表还是上一次的内容。")
    for warning in [*output.warnings, *output_warnings]:
        _warn(warning)
    console.print("[bold green]OK[/] 渲染完成")


# --------------------------------------------------------------------------- #
# validate
# --------------------------------------------------------------------------- #
@app.command("validate")
def validate_command(
    config: Path = typer.Option(
        ...,
        "--config",
        "-c",
        exists=True,
        dir_okay=False,
        readable=True,
        help="YAML 配置文件路径",
    ),
    excel: Optional[Path] = typer.Option(
        None, "--excel", "-x", help="同时校验填写的 Excel 模板是否存在且结构正确"
    ),
) -> None:
    """校验 YAML 配置（以及可选地校验 Excel 结构），不产生任何输出。"""
    warnings: list[str] = []
    try:
        project = load_config(config)
        environment = build_environment()
        # 派生参数：语法 / 引用范围 / 循环 / 能否翻译成 Excel 公式
        warnings.extend(derived_validate_config(project, env=environment))

        template_table = Table(title="模板清单", header_style="bold cyan")
        template_table.add_column("名称", style="bold")
        template_table.add_column("引擎")
        template_table.add_column("输出位置")
        template_table.add_column("导出文件名")
        template_table.add_column("引用变量")

        used_by_template: dict[str, set[str]] = {}
        formula_templates: list[str] = []
        for template in project.templates:
            validate_template(template, env=environment, base_dir=project.source_dir)
            used = collect_variables(template, env=environment, base_dir=project.source_dir)
            used_by_template[template.name] = used
            unknown = sorted(name for name in used if name not in project.defined_names)
            if unknown:
                warnings.append(
                    f"模板 {template.name!r} 引用了未在 YAML 中定义的变量: {', '.join(unknown)}"
                    "（若这些名字出现在 Excel 表里则仍然可用）"
                )
            if template.engine == "excel":
                formula_templates.append(template.name)
                # 提前编译一遍：超出"纯替换"子集的写法在这里就报出具体行
                compiled = compile_formulas(
                    template,
                    project,
                    case_columns=[FIRST_CASE_COLUMN],
                    source=template_source(template, project.source_dir),
                )
                longest = max(
                    (len(line) for case in compiled for line in case), default=0
                )
                if longest > LONG_FORMULA_WARN:
                    warnings.append(
                        f"模板 {template.name!r} 的最长公式 {longest} 字符"
                        f"（警告阈值 {LONG_FORMULA_WARN}）：一个 {{{{ x }}}} 约展开 300–400 字符，"
                        "一行超过 8 个占位符就该考虑拆行"
                    )
            template_table.add_row(
                template.name,
                "excel·公式" if template.engine == "excel" else "snapshot·快照",
                f"{template.output_sheet} @ {template.start_cell} ({template.direction})",
                template.filename or f"{template.name}_{{{{ case_name }}}}{template.extension}",
                ", ".join(sorted(used)) or "-",
            )
            if template.case_filter:
                console.print(
                    f"  [dim]{template.name}：case_filter = {template.case_filter}[/]"
                )
        console.print(template_table)

        if formula_templates:
            console.print(
                f"[cyan]i[/] 公式模式（engine: excel）的模板: {', '.join(formula_templates)} —— "
                "改参数后 Excel 打开即重算；"
                "但要注意：值与公式都只在 Excel/WPS 里成立，导出代码文件（--outdir）与 CI 检查（check）仍走命令行。"
            )

        variables_table = Table(title="变量清单", header_style="bold cyan")
        variables_table.add_column("作用域")
        variables_table.add_column("名称", style="bold")
        variables_table.add_column("来源")
        variables_table.add_column("类型")
        variables_table.add_column("默认值")
        variables_table.add_column("Prefix")
        variables_table.add_column("Suffix")
        variables_table.add_column("描述")
        for scope, items in (
            ("global", project.global_variables),
            ("local", project.local_variables),
        ):
            for variable in items:
                variables_table.add_row(
                    scope,
                    variable.name,
                    "计算" if variable.is_derived else "填写",
                    variable.type,
                    str(variable.default),
                    variable.prefix or "-",
                    variable.suffix or "-",
                    variable.description or "-",
                )
        console.print(variables_table)

        overlap = sorted(set(project.global_names) & set(project.local_names))
        if overlap:
            warnings.append(
                f"变量名同时出现在 global 与 local: {', '.join(overlap)}（local 会覆盖 global）"
            )

        # "定义了却没人用"是"改了参数没生效"的常见来源，单独提示
        # （派生表达式里引用到的名字也算"被用到"，否则每个派生链都会误报）
        referenced: set[str] = set()
        for names in used_by_template.values():
            referenced |= names
        for variable in [*project.global_variables, *project.local_variables]:
            if variable.is_derived:
                referenced |= expression_names(variable.derived or "", env=environment)
        unused = [
            name
            for name in [*project.global_names, *project.local_names]
            if name not in referenced
        ]
        if unused:
            warnings.append(
                f"YAML 定义了但没有被任何模板引用的变量: {', '.join(unused)}"
                "（拼写错误？或只是给 Excel 读者看的标注变量）"
            )

        if excel is not None:
            excel_path = Path(excel)
            workbook = load_workbook_file(excel_path)
            try:
                check_required_sheets(workbook, project)
                global_values = read_global_values(workbook, project)
                cases = read_cases(workbook, project)
                metadata = read_metadata(workbook, project)
            finally:
                workbook.close()
            # 取值约束：这一条会让"表里填错了一个数字"在 validate 阶段就暴露
            check_value_constraints(project, global_values, cases)
            excel_table = Table(title="Excel 检查", header_style="bold cyan")
            excel_table.add_column("项目", style="bold")
            excel_table.add_column("内容")
            excel_table.add_row("文件", str(excel_path))
            excel_table.add_row("Global 变量", ", ".join(global_values) or "-")
            excel_table.add_row(
                "Case 列",
                ", ".join(f"{case.name}({column_index_to_letter(case.column)})" for case in cases),
            )
            for case in cases:
                excel_table.add_row(
                    f"Case {case.name}",
                    ", ".join(f"{key}={value}" for key, value in case.values.items()) or "-",
                )
            if metadata:
                excel_table.add_row("上次渲染", metadata.get("时间", "-"))
                excel_table.add_row(
                    "参数指纹",
                    f"{metadata.get('参数指纹', '-')}"
                    f"（当前 {input_fingerprint(global_values, cases)}）",
                )
            else:
                excel_table.add_row("上次渲染", "（无记录：还没跑过 render --write-excel）")
            console.print(excel_table)
    except CodeGenError as exc:
        raise _fail(exc) from exc
    except (ValueError, OSError) as exc:
        raise _fail(exc) from exc

    for warning in warnings:
        _warn(warning)
    console.print("[bold green]OK[/] 配置校验通过")


# --------------------------------------------------------------------------- #
# check：输出表是否已过期
# --------------------------------------------------------------------------- #
def _last_non_empty_row(worksheet, start_row: int, column: int) -> int:
    last = start_row - 1
    for row in range(start_row, worksheet.max_row + 1):
        if worksheet.cell(row=row, column=column).value not in (None, ""):
            last = row
    return last


def _last_non_empty_column(worksheet, row: int, start_column: int) -> int:
    last = start_column - 1
    for column in range(start_column, worksheet.max_column + 1):
        if worksheet.cell(row=row, column=column).value not in (None, ""):
            last = column
    return last


def _read_column(worksheet, column: int, start_row: int) -> list[str]:
    last = _last_non_empty_row(worksheet, start_row, column)
    if last < start_row:
        return []
    return [to_text(worksheet.cell(row=row, column=column).value) for row in range(start_row, last + 1)]


def _read_row(worksheet, row: int, start_column: int) -> list[str]:
    last = _last_non_empty_column(worksheet, row, start_column)
    if last < start_column:
        return []
    return [to_text(worksheet.cell(row=row, column=column).value) for column in range(start_column, last + 1)]


def _first_difference(actual: list[str], expected: list[str]) -> str:
    for index in range(min(len(actual), len(expected))):
        if actual[index] != expected[index]:
            return f"第 {index + 1} 行不同：表里 {actual[index]!r}，应为 {expected[index]!r}"
    if len(actual) != len(expected):
        return f"行数不同：表里 {len(actual)} 行，应为 {len(expected)} 行"
    return "内容不同"


def _expected_lines(
    workbook,
    project: ProjectConfig,
    template,
    results: Sequence[RenderResult],
) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    """每个 Case 期望出现在输出表里的内容，以及"人读的标签"。

    快照模式 = 渲染文本；公式模式 = 编译出来的公式，标签是对应的模板源行
    （差异信息里贴模板行，比贴几百字符的公式有用得多）。

    :returns: ``(期望内容, 模板源行)``
    """
    if template.engine != "excel":
        return {result.case_name: result.lines for result in results}, {}

    columns = case_column_map(workbook, project)
    missing = [result.case_name for result in results if result.case_name not in columns]
    if missing:
        raise ExcelError(
            f"模板 {template.name!r} 使用公式模式，但 Local 表里找不到这些 Case 列: {', '.join(missing)}"
        )
    source = template_source(template, project.source_dir)
    per_case = compile_formulas(
        template,
        project,
        case_columns=[columns[result.case_name] for result in results],
        source=source,
    )
    labels = {case_name: source.splitlines() for case_name in (r.case_name for r in results)}
    return {result.case_name: lines for result, lines in zip(results, per_case)}, labels


def _describe_line(index: int, labels: Sequence[str] | None) -> str:
    if labels and index < len(labels):
        return f"（模板第 {index + 1} 行：{labels[index].strip()!r}）"
    return ""


def _first_difference(
    actual: list[str],
    expected: list[str],
    *,
    labels: Sequence[str] | None = None,
    what: str = "内容",
) -> str:
    for index in range(min(len(actual), len(expected))):
        if actual[index] != expected[index]:
            if labels is not None:
                return (
                    f"第 {index + 1} 行不同{_describe_line(index, labels)}；"
                    f"表里 {_clip(actual[index])}，应为 {_clip(expected[index])}"
                )
            return f"第 {index + 1} 行不同：表里 {actual[index]!r}，应为 {expected[index]!r}"
    if len(actual) != len(expected):
        return f"行数不同：表里 {len(actual)} 行，应为 {len(expected)} 行"
    return f"{what}不同"


def _clip(text: str, limit: int = 80) -> str:
    """把可能几百字符的公式裁到人能看的长度。"""
    shown = text if len(text) <= limit else text[:limit] + "…"
    return repr(shown)


def _output_differences(
    workbook,
    project: ProjectConfig,
    fresh: RenderOutput,
    *,
    verify_values: bool = True,
) -> tuple[list[str], list[str]]:
    """返回 ``(问题, 提示)``。

    * 快照模式：逐行比渲染文本。
    * 公式模式：① 比公式文本（过期 → 需要重跑 ``--write-excel``）；
      ② **把公式在 Python 里算一遍**，与 Python 渲染逐行比对 —— 这一条能抓到
      "列标指错 / 该用 ISBLANK 却用 =''" 这类生成端问题，也能抓到"参数表结构变了"。
    """
    problems: list[str] = []
    notes: list[str] = []
    for template in project.templates:
        results: list[RenderResult] = list(fresh.results.get(template.name) or [])
        if template.output_sheet not in workbook.sheetnames:
            problems.append(f"{template.name}: 工作表 {template.output_sheet!r} 不存在")
            continue
        worksheet = workbook[template.output_sheet]
        column, row = parse_cell(template.start_cell)
        if not results:
            problems.append(
                f"{template.name}: 当前没有任何 Case 匹配 case_filter，"
                f"{template.output_sheet} 表里的旧内容无法核对"
            )
            continue
        expected, labels = _expected_lines(workbook, project, template, results)

        evaluated: dict[str, list[str]] | None = None
        if template.engine == "excel" and verify_values:
            try:
                evaluated = evaluate_template_values(
                    workbook, project, template, [r.case_name for r in results]
                )
            except FormulaEvalError as exc:
                notes.append(f"{template.name}: 公式值校验已跳过 —— {exc}")

        for index, result in enumerate(results):
            if template.direction == "horizontal":
                at = column + index
                label = f"{template.output_sheet} 第 {column_index_to_letter(at)} 列（{result.case_name}）"
                actual = _read_column(worksheet, at, row)
                header = (
                    to_text(worksheet.cell(row=row - 1, column=at).value)
                    if template.write_case_headers and row > 1
                    else result.case_name
                )
            else:
                at = row + index
                label = f"{template.output_sheet} 第 {at} 行（{result.case_name}）"
                actual = _read_row(worksheet, at, column)
                header = (
                    to_text(worksheet.cell(row=at, column=column - 1).value)
                    if template.write_case_headers and column > 1
                    else result.case_name
                )
            if header != result.case_name:
                problems.append(
                    f"{label}：Case 表头是 {header!r}，应为 {result.case_name!r}"
                )
            want = expected[result.case_name]
            formulas_match = actual == want
            if not formulas_match:
                detail = _first_difference(
                    actual,
                    want,
                    labels=labels.get(result.case_name) if labels else None,
                )
                if template.engine == "excel":
                    detail += "（公式模式：比的是公式，重跑 --write-excel 刷新）"
                problems.append(f"{label}：{detail}")
            elif evaluated is not None:
                # 公式文本一致，再把公式算一遍与 Python 渲染对比
                got = evaluated.get(result.case_name, [])
                if got != result.lines:
                    problems.append(
                        f"{label}：`公式算出来的文本`与 Python 渲染不一致 → "
                        f"{_first_difference(got, result.lines)}"
                        "（参数表结构改动过？插/删过 Case 列？请重跑 --write-excel）"
                    )
    return problems, notes


@app.command("check")
def check_command(
    config: Path = typer.Option(
        ...,
        "--config",
        "-c",
        exists=True,
        dir_okay=False,
        readable=True,
        help="YAML 配置文件路径",
    ),
    excel: Optional[Path] = typer.Option(
        None, "--excel", "-x", help="要检查的 Excel（默认取配置中的 excel.output）"
    ),
    verify_values: bool = typer.Option(
        True,
        "--values/--no-values",
        help="公式模式：把 Output 表的公式在 Python 里算一遍，与 Python 渲染比对（推荐开）",
    ),
) -> None:
    """检查 Excel 里的输出表是否与当前参数一致；过期则退出码 1（可放进 CI）。"""
    try:
        project = load_config(config)
        excel_path = _open_excel(project, excel)
        fresh = render_all(project, excel_path)
        workbook = load_workbook_file(excel_path)
        try:
            recorded = read_metadata(workbook, project)
            problems, notes = _output_differences(
                workbook, project, fresh, verify_values=verify_values
            )
        finally:
            workbook.close()
        now_input = input_fingerprint(fresh.global_values, fresh.cases)
        now_output = output_fingerprint(fresh.results)
    except CodeGenError as exc:
        raise _fail(exc) from exc
    except (ValueError, OSError) as exc:
        raise _fail(exc) from exc

    for warning in fresh.warnings:
        _warn(warning)
    for note in notes:
        console.print(f"[yellow]![/] {note}")

    recorded_input = recorded.get("参数指纹", "")
    recorded_output = recorded.get("输出指纹", "")
    snapshot_names = [t.name for t in project.templates if t.engine != "excel"]
    formula_names = [t.name for t in project.templates if t.engine == "excel"]
    drift = bool(recorded_input) and recorded_input != now_input
    drift_note = ""
    if drift:
        drift_note = (
            "   ← 参数改过了（快照模板需要重跑）"
            if snapshot_names
            else "   ← 参数改过了；公式模板会自动重算，无需重跑"
        )

    table = Table(title="过期检查", header_style="bold cyan")
    table.add_column("项目", style="bold")
    table.add_column("内容")
    table.add_row("配置文件", str(config))
    table.add_row("Excel", str(excel_path))
    table.add_row("上次渲染时间", recorded.get("时间", "（无记录）"))
    table.add_row("模板引擎", "公式: " + (", ".join(formula_names) or "（无）")
                  + " ｜ 快照: " + (", ".join(snapshot_names) or "（无）"))
    table.add_row("参数指纹", f"记录 {recorded_input or '（无）'} / 当前 {now_input}" + drift_note)
    table.add_row("输出指纹", f"记录 {recorded_output or '（无）'} / 当前 {now_output}")
    console.print(table)

    if not recorded:
        _warn(
            "这个工作簿里没有渲染记录（没跑过 `render --write-excel`？）。"
            "下面按内容逐行比对。"
        )
    if formula_names and not snapshot_names:
        console.print(
            "[cyan]i[/] 全部模板都是公式模式：改参数不需要重跑；"
            "check 会同时比「公式是否与当前 YAML 一致」和「公式算出来的文本是否与 Python 渲染一致」。"
        )
    if formula_names and verify_values:
        console.print(
            "[cyan]i[/] 值校验已开启（--no-values 可关闭）：公式在 Python 里算了一遍再比对。"
        )
    if problems:
        error_console.print("[bold red]ERROR[/] 输出表已过期，共 "
                            f"{len(problems)} 处不一致：")
        for problem in problems:
            error_console.print(f"    {problem}")
        error_console.print(
            "    → 跑一次 `excel-codegen render -c "
            f"{config} -x {excel_path} --write-excel` 刷新"
        )
        raise typer.Exit(code=1)
    console.print("[bold green]OK[/] 输出表与当前参数一致")
