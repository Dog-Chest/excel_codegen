"""命令行接口：``spreadsheet-codegen init | render | validate | check``。

控制台注意事项：中文 Windows 的 GBK 控制台编不出 ``✓`` / ``✗`` 之类的字符，
所以本模块只用 ASCII 标记（``OK`` / ``ERROR`` / ``!``），并在导入时给
stdout/stderr 加上 ``errors="backslashreplace"`` 兜底 —— 成功路径绝不允许
因为"最后一行字打不出来"而返回非 0 退出码。
"""

from __future__ import annotations

import json
import sys
from collections.abc import Sequence
from contextlib import suppress
from pathlib import Path

import typer
from rich.console import Console
from rich.panel import Panel
from rich.syntax import Syntax
from rich.table import Table

from . import __version__, example_pack
from .derived import expression_names
from .derived import validate_config as derived_validate_config
from .excel_io import (
    case_axis_map,
    check_required_sheets,
    check_value_constraints,
    create_template,
    input_fingerprint,
    load_workbook_file,
    output_fingerprint,
    read_cases,
    read_global_values,
    read_group_members,
    read_metadata,
    template_source,
    workbook_deviates_from_defaults,
    write_results,
)
from .formula import LONG_FORMULA_WARN, FormulaError, compile_formulas, long_formula_warning, longest_formula
from .formula_eval import FormulaEvalError, evaluate_template_values
from .models import ProjectConfig, RenderResult, load_config
from .renderer import (
    RenderOutput,
    build_environment,
    check_asserts,
    collect_variables,
    compile_asserts,
    export_files,
    filename_collisions,
    plan_export_files,
    render_all,
    validate_template,
)
from .utils import (
    CodeGenError,
    ConfigError,
    ExcelError,
    InputError,
    WorkbookIOError,
    column_index_to_letter,
    parse_cell,
    to_text,
)


def _make_streams_forgiving() -> None:
    """让 stdout/stderr 遇到当前编码表达不了的字符时降级，而不是抛异常。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        # 宿主环境千奇百怪：不支持就静默跳过，别让"打日志"把命令搞崩
        with suppress(ValueError, OSError):  # pragma: no cover - 取决于宿主环境
            reconfigure(errors="backslashreplace")


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
#: 退出码分类（见 docs/cli.md）。以前所有失败都退 1，CI 里想区分"参数填错了"
#: 与"表里的输出过期了"只能 grep 文本 —— 而这两件事的处置完全不同。
EXIT_OK = 0
#: 通用失败（历史默认值）。保留它，老脚本里的 ``if errorlevel 1`` 照常能用。
EXIT_FAILURE = 1
#: 配置 / 取值错误：YAML 写错、变量越界、asserts 不满足 —— "你给的东西不对"。
EXIT_BAD_CONFIG = 2
#: 输出过期：``check`` 发现表里的内容与当前参数不一致 —— "该重跑一次了"。
EXIT_STALE = 3
#: 环境 / 依赖问题：工作簿打不开、文件被占用这类"不是配置的问题"。
EXIT_ENVIRONMENT = 4


def _fail(message: object, code: int = EXIT_FAILURE) -> typer.Exit:
    error_console.print(f"[bold red]ERROR[/] {message}")
    return typer.Exit(code=code)


def _failure_code(exc: BaseException) -> int:
    """按错误性质给退出码。

    * :class:`WorkbookIOError` -> **4**：工作簿读不出来 / 写不进去（被 Excel 占着、
      文件坏了、没权限）—— 这是环境问题，重试或人工介入，不是"改配置"；
    * :class:`ConfigError` / :class:`InputError` / :class:`FormulaError` -> **2**：
      YAML 写错、表里取值越界 / 不满足 ``asserts``、模板写法公式模式表达不了 ——
      都是"用户该改配置或输入"的信号；
    * 其余（变量缺失、工作簿结构不对）-> **1**，与历史行为一致。

    这样 CI 里就能把"我的参数填错了"和"表里的输出过期了（3）"分开处理（见 docs/cli.md）。
    """
    if isinstance(exc, WorkbookIOError):
        return EXIT_ENVIRONMENT
    if isinstance(exc, (ConfigError, InputError, FormulaError)):
        return EXIT_BAD_CONFIG
    return EXIT_FAILURE


def _warn(message: str) -> None:
    console.print(f"[bold yellow]![/] {message}")


def _load_project(config: Path):
    """加载配置，并把 ``extends`` 合并过程中的告警打出来。"""
    project = load_config(config)
    for warning in project.load_warnings:
        _warn(warning)
    return project


def _version_callback(value: bool) -> None:
    if value:
        console.print(f"spreadsheet-codegen [cyan]{__version__}[/]")
        raise typer.Exit()


def _resolve_excel_path(config: ProjectConfig, excel: Path | None) -> Path:
    """定死"工作簿在哪"的规则，避免同一个 YAML 里两套路径语义。

    * ``-x`` / ``--excel`` 给了就照用 —— 命令行给的路径**相对当前工作目录**（终端用户的直觉）；
    * 否则用 ``excel.output``，它**相对配置文件所在目录**解析 —— 与 ``template_file``
      一致（最小惊讶原则）。放在 ``rules/`` 子目录里的配置写 ``output: "../book.xlsx"``
      也能落在预期位置，而不是"跟着 CWD 跑"。

    .. versionchanged:: 0.11.0
        ``excel.output`` 以前相对**当前工作目录**解析（与 ``template_file`` 不一致，
        "文件跑哪去了"很难查）。现在与 ``template_file`` 对齐，并且**规范化**成绝对路径
        （``rules/../book.xlsx`` 里的 ``..`` 不再原样留着）；变更见 CHANGELOG。

    解析不存在的路径**不报错**：调用方（``init`` / ``doctor``）需要区分"不存在"与"非法"。
    """
    if excel is not None:
        return Path(excel)
    # .resolve() 顺带把 .. 折掉：摘要里再也不会出现 "rules\..\book.xlsx" 这种要人脑补的路径
    return config_path_within(config, config.excel.output).resolve()


def config_path_within(config: ProjectConfig, raw: str | Path) -> Path:
    """把配置里的相对路径按"相对配置文件所在目录"解析（绝对路径原样返回）。"""
    path = Path(raw)
    if path.is_absolute() or config.source_dir is None:
        return path
    return Path(config.source_dir) / path


def _excel_path_note(config: ProjectConfig, excel: Path | None) -> str:
    """给 ``init`` / ``doctor`` 的一行说明：这个路径是怎么来的。

    "文件跑到别处去了"是最费时间的排查，所以把规则直接打出来（而不是让人去翻文档）。
    """
    if excel is not None:
        return "（命令行指定的路径，相对当前工作目录）"
    if Path(config.excel.output).is_absolute():
        return "（配置里的 excel.output 是绝对路径）"
    if config.source_dir is None:  # pragma: no cover - 经过 load_config 就一定有
        return ""
    return f"（excel.output：相对配置文件所在目录 {config.source_dir}）"


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


def _prefill_output(
    project: ProjectConfig,
    target: Path,
    *,
    template_sheet: bool = True,
    howto_sheet: bool = True,
) -> str:
    """建完表顺手把输出也写一遍，返回"给用户看的一句话"。

    为什么要它：本工具的用法是"生成一次工作簿，之后就在 Excel 里干活"。公式模式下
    输出表里是活公式，所以**建完就能用**，不该再让人跑第二条命令。

    参数还是 YAML 默认值，可能暂时过不了取值约束 / ``asserts``（比如"舱高不能超过型深"，
    而默认值恰好不合）。这时**不能**让 ``init`` 失败 —— 骨架照常给，只是跳过预填并说明原因。
    """
    from .renderer import render_all

    try:
        output = render_all(project, target)
        # `init --no-howto` 时别把 HOWTO 表写回来（write_results 会照着配置建）
        write_results(
            target,
            project,
            output.results,
            command="spreadsheet-codegen init",
            update_howto=howto_sheet,
        )
        _drop_suppressed_sheets(project, target, template_sheet=template_sheet, howto_sheet=howto_sheet)
    except CodeGenError as exc:
        first = str(exc).strip().splitlines()[0]
        console.print(
            Panel(
                f"[yellow]跳过预填输出[/]：{first}\n"
                "（多半是默认值还没填全 / 过不了取值约束或 asserts —— 骨架已生成，"
                "填好参数后跑 render --write-excel 即可）",
                border_style="yellow",
            )
        )
        return "骨架（预填被跳过，见上方提示）"

    formulas = sum(1 for item in project.templates if item.engine == "excel")
    if formulas == len(project.templates):
        return "已写入公式 —— 打开 Excel 改参数即自动重算，不用再跑命令 ✓"
    if formulas == 0:
        return "已写入文本快照 —— 改参数后要重跑 render --write-excel"
    return f"已写入（其中 {formulas} 个模板是公式，其余是快照）"


def _drop_suppressed_sheets(
    project: ProjectConfig,
    target: Path,
    *,
    template_sheet: bool,
    howto_sheet: bool,
) -> None:
    """``init --no-template-sheet`` / ``--no-howto`` 时，把预填过程中建回来的表删掉。

    ``write_results`` 是照着**配置**记录元信息的（它不知道本次 ``init`` 关掉了哪张表），
    所以这里按命令行开关收尾 —— 否则"我不想生成说明表"会被预填偷偷推翻。
    """
    unwanted = []
    if not template_sheet and project.excel.template_sheet:
        unwanted.append(project.excel.template_sheet)
    if not howto_sheet and project.excel.howto_sheet:
        unwanted.append(project.excel.howto_sheet)
    if not unwanted:
        return
    workbook = load_workbook_file(target)
    try:
        for name in unwanted:
            if name in workbook.sheetnames:
                workbook.remove(workbook[name])
        workbook.save(target)
    finally:
        workbook.close()


def _case_axis_label(project: ProjectConfig) -> str:
    """工况轴的说法：横向是"Case 列"，纵向是"Case 行"。"""
    return "Case 列" if project.excel.local_direction == "horizontal" else "Case 行"


def _local_layout_hint(project: ProjectConfig) -> str:
    if project.excel.local_direction == "horizontal":
        return "Case 列：E 起"
    return "Case 行：第 2 行起"


def _case_names_of(spec: int | list[str]) -> list[str]:
    return [f"Case{index}" for index in range(1, spec + 1)] if isinstance(spec, int) else list(spec)


def _open_excel(project: ProjectConfig, excel: Path | None) -> Path:
    path = _resolve_excel_path(project, excel)
    if not path.exists():
        raise ExcelError(f"Excel 文件不存在: {path}（请先运行 `spreadsheet-codegen init` 生成模板并填写参数）")
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
    output: Path | None = typer.Option(
        None, "--output", "-o", help="输出的 Excel 模板路径（默认取配置中的 excel.output）"
    ),
    cases: str = typer.Option(
        "2",
        "--cases",
        help="初始 Case 列：数量（如 3）或逗号分隔的名字（如 EXT-T20,INT-T15）",
    ),
    force: bool = typer.Option(False, "--force", "-f", help="目标 Excel 已存在时覆盖"),
    yes: bool = typer.Option(
        False,
        "--yes",
        "-y",
        help="跳过「这本工作簿里填过参数」的确认（配合 --force：会丢掉已填的参数）",
    ),
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
    comments: bool = typer.Option(
        True,
        "--comments/--no-comments",
        help="是否给变量名那格加 Excel 批注（描述 / 单位 / 约束 / 前缀后缀 / 派生表达式）",
    ),
    scripts: bool = typer.Option(
        True,
        "--scripts/--no-scripts",
        help="是否在工作簿旁边生成一键刷新脚本（<工作簿名>_render.bat / .sh）",
    ),
    prerender: bool = typer.Option(
        True,
        "--prerender/--no-prerender",
        help="建完表就先把输出写进去（默认打开：公式模式下打开工作簿即可用）",
    ),
) -> None:
    """根据 YAML 生成 Excel 参数填写模板。

    默认（``--prerender``）建完骨架就把输出也写一遍 —— 公式模式下输出表里是活公式，
    于是"生成完就能打开 Excel 干活"，不需要再跑一条命令。参数没填全、
    暂时过不了约束/asserts 时会跳过这一步并提示（骨架照常生成）。
    """
    try:
        spec = _parse_cases(cases)
        project = _load_project(config)
        target = output if output is not None else _resolve_excel_path(project, None)
        # --force 是**数据丢失点**：重建会盖掉已填的参数。所以"人填过东西"时要显式 --yes。
        # 判据是参数指纹（等于全默认值的指纹 = 没人动过），所以刚生成的工作簿不会被拦。
        if force and not yes and Path(target).exists():
            note = workbook_deviates_from_defaults(target, project)
            if note is not None:
                raise _fail(
                    f"{note}。\n"
                    f"  重建会把这些参数换成 YAML 默认值（原文件不会自动备份）。\n"
                    f"  → 确认要丢掉就加 --yes；只想改骨架不想丢数据，就别加 --force",
                    code=EXIT_BAD_CONFIG,
                )
        target = create_template(
            project,
            target,
            cases=spec,
            overwrite=force,
            include_template_sheet=template_sheet,
            include_howto_sheet=howto_sheet,
            include_comments=comments,
            include_scripts=scripts,
        )
    except CodeGenError as exc:
        raise _fail(exc, code=_failure_code(exc)) from exc
    except OSError as exc:
        # 打不开 / 写不进工作簿：不是配置问题，单独一个退出码
        raise _fail(exc, code=EXIT_ENVIRONMENT) from exc
    except ValueError as exc:
        raise _fail(exc) from exc

    prefill_note = "（还没有写输出）"
    if prerender:
        prefill_note = _prefill_output(project, target, template_sheet=template_sheet, howto_sheet=howto_sheet)

    table = Table(title="Excel 模板已生成", show_header=True, header_style="bold cyan")
    table.add_column("项目", style="bold")
    # 同"渲染摘要"：路径折行，不许被截断成 "…"
    table.add_column("内容", no_wrap=False, overflow="fold")
    # 绝对路径：同一个 YAML 里 template_file 相对 YAML、excel.output 相对 YAML ——
    # 规则统一之后，"到底写哪去了"也一并写死，省掉一轮"文件不见了"的排查
    table.add_row("文件", f"{target.resolve()}\n[dim]{_excel_path_note(project, output)}[/]")
    table.add_row("Global 表", project.excel.sheets.global_)
    table.add_row("Local 表", f"{project.excel.sheets.local}（{_local_layout_hint(project)}）")
    table.add_row("Output 表", ", ".join(project.excel.sheets.outputs))
    table.add_row(_case_axis_label(project), ", ".join(_case_names_of(spec)))
    table.add_row("说明表", project.excel.howto_sheet or "（未生成）")
    table.add_row("模板数", str(len(project.templates)))
    table.add_row("输出内容", prefill_note)
    console.print(table)
    # 下一步的命令里写**规范化后的绝对路径**（Windows 上反斜杠）：可以直接粘进 Explorer / Excel，
    # 也免得相对路径在别的目录下跑时指错文件
    absolute_target = str(target.resolve())
    console.print(
        Panel(
            "[bold]下一步[/]\n"
            f"1. 在 [cyan]{project.excel.sheets.global_}[/] 表填写 B 列（Value），D/E 列可覆盖 Prefix/Suffix；\n"
            f"2. 在 [cyan]{project.excel.sheets.local}[/] 表从 E 列开始按 Case 填写，如需更多 Case 直接右拉复制；\n"
            "3. 运行渲染（不加 --write-excel 只预览，不会改动 Excel）：\n"
            f'   [green]spreadsheet-codegen render --config "{config}" --excel "{absolute_target}" --write-excel[/]\n'
            f'   [green]spreadsheet-codegen render --config "{config}" --excel "{absolute_target}" --outdir generated/[/]',
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
    excel: Path | None = typer.Option(None, "--excel", "-x", help="已填写的 Excel 模板（默认取配置中的 excel.output）"),
    write_excel: bool = typer.Option(False, "--write-excel", "-w", help="把渲染结果写入 Excel 的 Output 表"),
    outdir: Path | None = typer.Option(None, "--outdir", "-d", help="把渲染结果导出为代码文件到该目录"),
    case: list[str] | None = typer.Option(None, "--case", help="只渲染指定 Case（可重复传入）"),
    show: bool = typer.Option(True, "--show/--no-show", help="在终端打印渲染结果"),
    overwrite: bool = typer.Option(True, "--overwrite/--no-overwrite", help="导出文件已存在时是否覆盖"),
    allow_overwrite_filename: bool = typer.Option(
        False,
        "--allow-overwrite-filename",
        help="放行「多份结果写进同一个文件名」（默认拦住：那会静默只留下最后一份）",
    ),
) -> None:
    """读取填好的 Excel + YAML，渲染模板并输出到 Excel / 代码文件 / 终端。

    ``--write-excel`` 与 ``--outdir`` 可以一起给：同一次渲染既刷新 Excel 的 Output 表、
    又把代码文件导出到目录 —— 改了参数想把两边都同步时不必跑两遍。
    """
    try:
        project = _load_project(config)
        excel_path = _open_excel(project, excel)
        output = render_all(project, excel_path, only_cases=list(case) if case else None)

        written_files: list[Path] = []
        collision_warnings: list[str] = []
        output_warnings: list[str] = []
        command_text = (
            f"spreadsheet-codegen render -c {config} -x {excel_path}"
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
            # 文件名冲突先算出来：--allow-overwrite-filename 时它是**告警**（说明最终留下的是哪一份）
            planned = plan_export_files(project, output.results, outdir)
            if allow_overwrite_filename:
                problem = filename_collisions(planned, outdir)
                if problem is not None:
                    collision_warnings.append(
                        "多份结果写进了同一个文件（--allow-overwrite-filename 放行）：\n" + problem
                    )
            written_files = export_files(
                project,
                output.results,
                outdir,
                overwrite=overwrite,
                allow_collisions=allow_overwrite_filename,
            )
    except CodeGenError as exc:
        raise _fail(exc, code=_failure_code(exc)) from exc
    except OSError as exc:
        # 打不开 / 写不进工作簿：不是配置问题，单独一个退出码
        raise _fail(exc, code=EXIT_ENVIRONMENT) from exc
    except ValueError as exc:
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
    # 内容列不省略（ellipsis=False）：长路径会**折行**而不是被截成 "…" ——
    # 摘要里"路径只显示一半"等于没显示，用户还是不知道文件写到哪去了。
    summary.add_column("内容", no_wrap=False, overflow="fold")
    summary.add_row("配置文件", str(config))
    summary.add_row("Excel", str(excel_path))
    summary.add_row("Case", ", ".join(item.name for item in output.cases))
    summary.add_row("渲染结果", f"{output.total()} 个（{len(project.templates)} 模板 × {len(output.cases)} Case）")
    if output.skipped:
        detail = "；".join(f"{name} 跳过 {', '.join(cases)}" for name, cases in output.skipped.items())
        summary.add_row("case_filter", f"跳过 {output.skipped_total()} 个（{detail}）")
    # 永远显示这一行：让"成功"与"成功但没动文件"能一眼分开
    summary.add_row("写回 Excel", "是" if write_excel else "否（需要 --write-excel）")
    if outdir is not None:
        # 绝对路径：路径语义一旦有歧义，"文件跑哪去了"最费时间（相对 YAML / 相对 CWD 之争）
        summary.add_row("导出目录", str(Path(outdir).resolve()))
    summary.add_row("导出文件", "\n".join(str(item) for item in written_files) if written_files else "无")
    console.print(summary)

    if not write_excel and outdir is None:
        _warn("本次只预览：没有写回 Excel，也没有导出文件。加 --write-excel / --outdir 才会落盘。")
    elif not write_excel:
        _warn(f"没有写回 Excel（只导出了文件）：Excel 里的 {project.excel.sheets.outputs[0]} 表还是上一次的内容。")
    for warning in [*output.warnings, *output_warnings, *collision_warnings]:
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
    excel: Path | None = typer.Option(None, "--excel", "-x", help="同时校验填写的 Excel 模板是否存在且结构正确"),
) -> None:
    """校验 YAML 配置（以及可选地校验 Excel 结构），不产生任何输出。"""
    warnings: list[str] = []
    try:
        project = _load_project(config)
        environment = build_environment()
        # 派生参数：语法 / 引用范围 / 循环 / 能否翻译成 Excel 公式
        warnings.extend(derived_validate_config(project, env=environment))
        # 跨变量校验：表达式先编译一遍（语法错误在这里就报，别等渲染）
        asserts = compile_asserts(project, env=environment)
        if asserts:
            console.print(f"  [dim]asserts（跨变量校验）{len(asserts)} 条：[/]")
            for expression, _ in asserts:
                console.print(f"      {expression}")

        template_table = Table(title="模板清单", header_style="bold cyan")
        template_table.add_column("名称", style="bold")
        template_table.add_column("引擎")
        template_table.add_column("输出位置")
        template_table.add_column("导出文件名")
        template_table.add_column("引用变量")

        used_by_template: dict[str, set[str]] = {}
        formula_templates: list[str] = []
        for template in project.templates:
            # 不传 env：让 renderer 按『声明这个模板的目录』建环境（{% include %} 要用）
            validate_template(template, base_dir=project.source_dir)
            used = collect_variables(template, base_dir=project.source_dir)
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
                source = template_source(template, project.source_dir)
                compiled = compile_formulas(
                    template,
                    project,
                    case_axes=[1],  # 只量长度，轴取哪个都行
                    source=source,
                )
                longest, placeholders = longest_formula(compiled, source)
                if longest > LONG_FORMULA_WARN:
                    warnings.append(long_formula_warning(template.name, longest, placeholders=placeholders))
            template_table.add_row(
                template.name,
                "excel·公式" if template.engine == "excel" else "snapshot·快照",
                f"{template.output_sheet} @ {template.start_cell} ({template.direction})",
                template.filename or f"{template.name}_{{{{ case_name }}}}{template.extension}",
                ", ".join(sorted(used)) or "-",
            )
            if template.case_filter:
                console.print(f"  [dim]{template.name}：case_filter = {template.case_filter}[/]")
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
        # default 的两种用途分开写：前者 = init 预填进新表，后者 = 空单元格回落。
        # 不写清就会踩"某类型才有的字段被默认值污染"（指南 §3.3）
        variables_table.add_column("取值")
        variables_table.add_column("Prefix")
        variables_table.add_column("Suffix")
        variables_table.add_column("描述")
        for scope, items in (
            ("global", project.global_variables),
            ("local", project.local_variables),
        ):
            for variable in items:
                if variable.is_derived:
                    behavior = "算式"
                elif variable.prefill and variable.fallback:
                    behavior = "预填 + 空则回落"
                elif variable.prefill:
                    behavior = "只预填"
                elif variable.fallback:
                    behavior = "只回落"
                else:
                    behavior = "不预填、不回落"
                if variable.allow_blank:
                    behavior += "；可留空"
                variables_table.add_row(
                    scope,
                    variable.name,
                    "计算" if variable.is_derived else "填写",
                    variable.type,
                    str(variable.default),
                    behavior,
                    variable.prefix or "-",
                    variable.suffix or "-",
                    variable.description or "-",
                )
        console.print(variables_table)

        overlap = sorted(set(project.global_names) & set(project.local_names))
        if overlap:
            warnings.append(f"变量名同时出现在 global 与 local: {', '.join(overlap)}（local 会覆盖 global）")

        # "定义了却没人用"是"改了参数没生效"的常见来源，单独提示
        # （派生表达式里引用到的名字也算"被用到"，否则每个派生链都会误报）
        referenced: set[str] = set()
        for names in used_by_template.values():
            referenced |= names
        for variable in [*project.global_variables, *project.local_variables]:
            if variable.is_derived:
                referenced |= expression_names(variable.derived or "", env=environment)
        unused = [name for name in [*project.global_names, *project.local_names] if name not in referenced]
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
                members = read_group_members(workbook, project)
                metadata = read_metadata(workbook, project)
            finally:
                workbook.close()
            # 取值约束：这一条会让"表里填错了一个数字"在 validate 阶段就暴露
            check_value_constraints(project, global_values, cases, members=members, warnings=warnings)
            check_asserts(project, global_values, cases, members=members, env=environment)
            excel_table = Table(title="Excel 检查", header_style="bold cyan")
            excel_table.add_column("项目", style="bold")
            excel_table.add_column("内容")
            excel_table.add_row("文件", str(excel_path))
            excel_table.add_row("Global 变量", ", ".join(global_values) or "-")
            excel_table.add_row(
                _case_axis_label(project),
                ", ".join(f"{case.name}({case.where})" for case in cases),
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
                    f"{metadata.get('参数指纹', '-')}（当前 {input_fingerprint(global_values, cases)}）",
                )
            else:
                excel_table.add_row("上次渲染", "（无记录：还没跑过 render --write-excel）")
            console.print(excel_table)
    except CodeGenError as exc:
        raise _fail(exc, code=_failure_code(exc)) from exc
    except OSError as exc:
        # 打不开 / 写不进工作簿：不是配置问题，单独一个退出码
        raise _fail(exc, code=EXIT_ENVIRONMENT) from exc
    except ValueError as exc:
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

    columns = case_axis_map(workbook, project)
    missing = [result.case_name for result in results if result.case_name not in columns]
    if missing:
        raise ExcelError(f"模板 {template.name!r} 使用公式模式，但 Local 表里找不到这些 Case 列: {', '.join(missing)}")
    source = template_source(template, project.source_dir)
    per_case = compile_formulas(
        template,
        project,
        case_axes=[columns[result.case_name] for result in results],
        source=source,
    )
    labels = {case_name: source.splitlines() for case_name in (r.case_name for r in results)}
    return {result.case_name: lines for result, lines in zip(results, per_case, strict=False)}, labels


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
    differences = _differences(actual, expected, labels=labels, what=what, limit=1)
    return differences[0] if differences else f"{what}不同"


def _differences(
    actual: list[str],
    expected: list[str],
    *,
    labels: Sequence[str] | None = None,
    what: str = "内容",
    limit: int = 5,
) -> list[str]:
    """列出**所有**不同的行（最多 ``limit`` 条，其余折成一句"还有 N 行"）。

    只报第一处差异在 CI 里很难定位 —— 用户想知道"一共差多少、都差在哪"。
    """
    found: list[str] = []
    for index in range(min(len(actual), len(expected))):
        if actual[index] == expected[index]:
            continue
        if len(found) < limit:
            if labels is not None:
                found.append(
                    f"第 {index + 1} 行不同{_describe_line(index, labels)}；"
                    f"表里 {_clip(actual[index])}，应为 {_clip(expected[index])}"
                )
            else:
                found.append(f"第 {index + 1} 行不同：表里 {actual[index]!r}，应为 {expected[index]!r}")
    hidden = sum(1 for index in range(min(len(actual), len(expected))) if actual[index] != expected[index]) - len(found)
    if hidden > 0:
        found.append(f"……还有 {hidden} 行不同（只列了前 {limit} 行）")
    if not found and len(actual) != len(expected):
        found.append(f"行数不同：表里 {len(actual)} 行，应为 {len(expected)} 行")
    if not found:
        found.append(f"{what}不同")
    return found


def _clip(text: str, limit: int = 80) -> str:
    """把可能几百字符的公式裁到人能看的长度。"""
    shown = text if len(text) <= limit else text[:limit] + "…"
    return repr(shown)


#: ``check`` 的问题分类：**稳定枚举**，给 CI / 看板按性质分流用（``check --json`` 的
#: ``problems[].kind``）。改这里的取值属于破坏性变更，要进 CHANGELOG。
_MISSING_SHEET = "missing_sheet"  #: 输出表不存在
_NO_MATCHING_CASE = "no_matching_case"  #: case_filter 把所有 Case 都跳过了，没法核对
_CASE_HEADER = "case_header"  #: 输出表的 Case 表头与当前参数表对不上
_OUTPUT_STALE = "output_stale"  #: 输出表内容与当前 YAML / 参数不一致 —— 重跑 --write-excel
_VALUE_MISMATCH = "value_mismatch"  #: 公式文本一致，但公式算出来的文本与 Python 渲染不同

#: ``(分类, 说明)``。分类见上面的常量。
_Problem = tuple[str, str]


def _output_differences(
    workbook,
    project: ProjectConfig,
    fresh: RenderOutput,
    *,
    verify_values: bool = True,
) -> tuple[list[_Problem], list[str]]:
    """返回 ``(问题, 提示)``；每个问题带一个**稳定的 kind**（给 CI 分类用，见下面的常量）。

    * 快照模式：逐行比渲染文本。
    * 公式模式：① 比公式文本（过期 → 需要重跑 ``--write-excel``）；
      ② **把公式在 Python 里算一遍**，与 Python 渲染逐行比对 —— 这一条能抓到
      "列标指错 / 该用 ISBLANK 却用 =''" 这类生成端问题，也能抓到"参数表结构变了"。
    """
    problems: list[_Problem] = []
    notes: list[str] = []
    for template in project.templates:
        results: list[RenderResult] = list(fresh.results.get(template.name) or [])
        if template.output_sheet not in workbook.sheetnames:
            problems.append((_MISSING_SHEET, f"{template.name}: 工作表 {template.output_sheet!r} 不存在"))
            continue
        worksheet = workbook[template.output_sheet]
        column, row = parse_cell(template.start_cell)
        if not results:
            problems.append(
                (
                    _NO_MATCHING_CASE,
                    f"{template.name}: 当前没有任何 Case 匹配 case_filter，"
                    f"{template.output_sheet} 表里的旧内容无法核对",
                )
            )
            continue
        expected, labels = _expected_lines(workbook, project, template, results)

        evaluated: dict[str, list[str]] | None = None
        if template.engine == "excel" and verify_values:
            try:
                evaluated = evaluate_template_values(workbook, project, template, [r.case_name for r in results])
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
                problems.append((_CASE_HEADER, f"{label}：Case 表头是 {header!r}，应为 {result.case_name!r}"))
            want = expected[result.case_name]
            formulas_match = actual == want
            if not formulas_match:
                detail = "；".join(
                    _differences(
                        actual,
                        want,
                        labels=labels.get(result.case_name) if labels else None,
                    )
                )
                if template.engine == "excel":
                    detail += "（公式模式：比的是公式，重跑 --write-excel 刷新）"
                problems.append((_OUTPUT_STALE, f"{label}：{detail}"))
            elif evaluated is not None:
                # 公式文本一致，再把公式算一遍与 Python 渲染对比
                got = evaluated.get(result.case_name, [])
                if got != result.lines:
                    problems.append(
                        (
                            _VALUE_MISMATCH,
                            f"{label}：`公式算出来的文本`与 Python 渲染不一致 → "
                            + "；".join(_differences(got, result.lines))
                            + f"（参数表结构改动过？插/删过 {_case_axis_label(project)}？请重跑 --write-excel）",
                        )
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
    excel: Path | None = typer.Option(None, "--excel", "-x", help="要检查的 Excel（默认取配置中的 excel.output）"),
    verify_values: bool = typer.Option(
        True,
        "--values/--no-values",
        help="公式模式：把 Output 表的公式在 Python 里算一遍，与 Python 渲染比对（推荐开）",
    ),
    json_output: bool = typer.Option(
        False,
        "--json",
        help="把结果打成 JSON 输出（给 CI / 看板消费），不打表格",
    ),
) -> None:
    """检查 Excel 里的输出表是否与当前参数一致。

    退出码：**3** = 输出过期（可放进 CI，与"配置写错"区分开）、
    2 = 配置/取值不对、4 = 环境问题（打不开工作簿）、0 = 一致。
    详见 docs/cli.md 的「退出码」一节。
    """
    try:
        project = _load_project(config)
        excel_path = _open_excel(project, excel)
        fresh = render_all(project, excel_path)
        workbook = load_workbook_file(excel_path)
        try:
            recorded = read_metadata(workbook, project)
            problems, notes = _output_differences(workbook, project, fresh, verify_values=verify_values)
        finally:
            workbook.close()
        now_input = input_fingerprint(fresh.global_values, fresh.cases)
        now_output = output_fingerprint(fresh.results)
    except CodeGenError as exc:
        raise _fail(exc, code=_failure_code(exc)) from exc
    except OSError as exc:
        # 打不开 / 写不进工作簿：不是配置问题，单独一个退出码
        raise _fail(exc, code=EXIT_ENVIRONMENT) from exc
    except ValueError as exc:
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
            "   ← 参数改过了（快照模板需要重跑）" if snapshot_names else "   ← 参数改过了；公式模板会自动重算，无需重跑"
        )

    if json_output:
        # --json：stdout 上只留 JSON，方便直接喂给 jq / CI 看板
        _print_check_json(
            config=config,
            excel_path=excel_path,
            recorded=recorded,
            now_input=now_input,
            now_output=now_output,
            warnings=[*fresh.warnings, *notes],
            problems=problems,
        )
        if problems:
            raise typer.Exit(code=EXIT_STALE)
        return

    table = Table(title="过期检查", header_style="bold cyan")
    table.add_column("项目", style="bold")
    table.add_column("内容")
    table.add_row("配置文件", str(config))
    table.add_row("Excel", str(excel_path))
    table.add_row("上次渲染时间", recorded.get("时间", "（无记录）"))
    table.add_row(
        "模板引擎",
        "公式: " + (", ".join(formula_names) or "（无）") + " ｜ 快照: " + (", ".join(snapshot_names) or "（无）"),
    )
    table.add_row("参数指纹", f"记录 {recorded_input or '（无）'} / 当前 {now_input}" + drift_note)
    table.add_row("输出指纹", f"记录 {recorded_output or '（无）'} / 当前 {now_output}")
    console.print(table)

    if not recorded:
        _warn("这个工作簿里没有渲染记录（没跑过 `render --write-excel`？）。下面按内容逐行比对。")
    if formula_names and not snapshot_names:
        console.print(
            "[cyan]i[/] 全部模板都是公式模式：改参数不需要重跑；"
            "check 会同时比「公式是否与当前 YAML 一致」和「公式算出来的文本是否与 Python 渲染一致」。"
        )
    if formula_names and verify_values:
        console.print("[cyan]i[/] 值校验已开启（--no-values 可关闭）：公式在 Python 里算了一遍再比对。")
    if problems:
        error_console.print(f"[bold red]ERROR[/] 输出表已过期，共 {len(problems)} 处不一致：")
        for _, problem in problems:
            error_console.print(f"    {problem}")
        error_console.print(f"    → 跑一次 `spreadsheet-codegen render -c {config} -x {excel_path} --write-excel` 刷新")
        raise typer.Exit(code=EXIT_STALE)
    console.print("[bold green]OK[/] 输出表与当前参数一致")


def _print_check_json(
    *,
    config: Path,
    excel_path: Path,
    recorded: dict[str, str],
    now_input: str,
    now_output: str,
    warnings: Sequence[str],
    problems: Sequence[_Problem],
) -> None:
    """``check --json``：给 CI / 看板消费的机读结果。

    退出码与表格模式一致（过期 = ``3``），所以两种模式可以互换。

    ``problems`` 保留**字符串数组**（老消费方不用改），同时新增 ``problem_kinds``
    给出稳定的分类枚举 —— CI 想按性质分流不必再去 grep 中文文本。
    """
    payload = {
        "ok": not problems,
        "config": str(config),
        "excel": str(excel_path),
        "recorded": {
            "time": recorded.get("时间", ""),
            "input_fingerprint": recorded.get("参数指纹", ""),
            "output_fingerprint": recorded.get("输出指纹", ""),
        },
        "current": {"input_fingerprint": now_input, "output_fingerprint": now_output},
        "drift": bool(recorded.get("参数指纹")) and recorded.get("参数指纹") != now_input,
        "warnings": list(warnings),
        # 两个并行的数组：problems 是给人的话术，problem_kinds 是给机器的分类
        "problems": [message for _, message in problems],
        "problem_kinds": [kind for kind, _ in problems],
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))


# --------------------------------------------------------------------------- #
# doctor：一条命令体检环境 / 配置 / 工作簿
# --------------------------------------------------------------------------- #
#: doctor 的一行结论：级别（OK / ! / ERROR）、项目、说明
_Finding = tuple[str, str, str]
#: doctor 收集到的一条错误：**退出码** + 提示原文。退出码要跟着异常一起留下来，
#: 否则汇总成报告后就分不清"配置错（2）"与"工作簿结构不对（1）"了。
_DoctorError = tuple[int, str]

#: 支持的 Python 下限（与 pyproject.toml 的 requires-python 对应）
_MIN_PYTHON = (3, 11)

#: 运行时依赖（与 pyproject.toml 的 [project.dependencies] 对应）
_RUNTIME_DEPS = ("openpyxl", "jinja2", "yaml", "pydantic", "typer", "rich")


def _doctor_environment() -> list[_Finding]:
    """环境这一层：Python 版本、依赖、可选工具、装法。"""
    import importlib.metadata as md
    import shutil

    findings: list[_Finding] = []

    version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    # 与 pyproject.toml 的 requires-python 保持一致。写成变量是故意的：
    # doctor 的职责就是"报告环境"，不能让 linter 把这个检查当成恒真优化掉。
    minimum = _MIN_PYTHON
    if sys.version_info[:2] >= minimum:
        findings.append(("OK", "Python", f"{version}（本项目要求 >={minimum[0]}.{minimum[1]}）"))
    else:
        findings.append(("ERROR", "Python", f"{version} 低于要求的 {minimum[0]}.{minimum[1]} —— 请升级解释器"))

    missing: list[str] = []
    versions: list[str] = []
    for name in _RUNTIME_DEPS:
        try:
            dist = "pyyaml" if name == "yaml" else name
            versions.append(f"{name} {md.version(dist)}")
        except md.PackageNotFoundError:
            missing.append(name)
    if missing:
        findings.append(("ERROR", "依赖", f"缺这些包：{', '.join(missing)}（跑一次 uv sync 或 pip install -e .）"))
    else:
        findings.append(("OK", "依赖", "、".join(versions)))

    if shutil.which("uv"):
        findings.append(("OK", "uv", "在（推荐用 uv run …，环境自洽）"))
    else:
        findings.append(("!", "uv", "没找到 —— 用 uv 可以免去 venv / pip 的差异（见 docs/setup.md）"))

    if shutil.which("git"):
        findings.append(("OK", "git", "在"))
    else:
        findings.append(("!", "git", "没找到 —— 建议用版本管理管住 YAML 与工作簿"))

    if sys.prefix != getattr(sys, "base_prefix", sys.prefix):
        findings.append(("OK", "运行环境", f"虚拟环境（{sys.prefix}）"))
    else:
        findings.append(("!", "运行环境", "用的是系统 Python —— 建议装在 venv / uv 环境里"))
    return findings


def _doctor_config(project: ProjectConfig, excel_path: Path | None) -> tuple[list[_Finding], list[_DoctorError]]:
    """配置这一层：结构统计 + 模板/派生/断言检查。返回 (结论, 错误列表)。

    错误列表带**退出码**（不是纯字符串）：``doctor`` 要把"配置错（2）"与
    "工作簿结构不对（1）"分开报出去，见 :func:`doctor_command`。
    """
    findings: list[_Finding] = []
    errors: list[_DoctorError] = []

    origin = f"extends {len(project.extends)} 个文件" if project.extends else "单文件"
    findings.append(("OK", "YAML", f"加载成功（{origin}）"))
    formula = [t.name for t in project.templates if t.engine == "excel"]
    snapshot = [t.name for t in project.templates if t.engine != "excel"]
    findings.append(
        (
            "OK",
            "规模",
            f"{len(project.templates)} 个模板（公式 {len(formula)} / 快照 {len(snapshot)}）、"
            f"{len(project.global_variables)} 个全局变量、{len(project.local_variables)} 个局部变量",
        )
    )

    constrained = [v.name for v in (*project.global_variables, *project.local_variables) if v.has_constraints]
    if constrained or project.asserts:
        findings.append(
            (
                "OK",
                "校验规则",
                f"{len(constrained)} 个变量带取值约束、{len(project.asserts)} 条跨变量 asserts",
            )
        )
    else:
        findings.append(("!", "校验规则", "一条约束都没有 —— 填错值不会被拦住（指南 §3.5 / §3.6）"))

    environment = build_environment()
    used_all: set[str] = set()
    for template in project.templates:
        try:
            validate_template(template, base_dir=project.source_dir)
            used_all |= collect_variables(template, base_dir=project.source_dir)
        except CodeGenError as exc:
            errors.append((_failure_code(exc), str(exc)))
            findings.append(("ERROR", f"模板 {template.name}", str(exc)))
            continue
        if template.engine == "excel":
            try:
                source = template_source(template, project.source_dir)
                compiled = compile_formulas(
                    template,
                    project,
                    case_axes=[1],  # 只量长度，轴取哪个都行
                    source=source,
                )
                longest, placeholders = longest_formula(compiled, source)
                if longest > LONG_FORMULA_WARN:
                    findings.append(
                        (
                            "!",
                            f"模板 {template.name}",
                            long_formula_warning(template.name, longest, placeholders=placeholders),
                        )
                    )
            except CodeGenError as exc:
                errors.append((_failure_code(exc), str(exc)))
                findings.append(("ERROR", f"模板 {template.name}", str(exc)))

    try:
        warnings = derived_validate_config(project, env=environment)
        findings.append(
            ("!" if warnings else "OK", "派生参数", "；".join(warnings) if warnings else "没有派生参数问题")
        )
    except CodeGenError as exc:
        errors.append((_failure_code(exc), str(exc)))
        findings.append(("ERROR", "派生参数", str(exc)))

    try:
        compile_asserts(project, env=environment)
    except CodeGenError as exc:
        errors.append((_failure_code(exc), str(exc)))
        findings.append(("ERROR", "asserts", str(exc)))

    # 只看 YAML 里真的定义了的变量：defined_names 含保留名（case_name / template_name），
    # 那是"模板里可以引用的名字"，不是"定义了没人用"
    declared = [*project.global_names, *project.local_names]
    unused = sorted(name for name in declared if name not in used_all)
    if unused:
        findings.append(("!", "未使用的变量", f"{', '.join(unused)} —— 定义了但没有模板引用（拼写错误？）"))
    else:
        findings.append(("OK", "变量使用", "YAML 里定义的变量都被模板用到了"))
    return findings, errors


def _doctor_workbook(project: ProjectConfig, excel_path: Path) -> tuple[list[_Finding], list[_DoctorError]]:
    """工作簿这一层：结构、取值、指纹。返回 (结论, 错误列表)。

    "打不开"与"读不动"都必须变成报告里的一行结论，**而不是 traceback** —— 用户跑
    ``doctor`` 的时候，工作簿往往正好是坏的（被 Excel 占着、写坏了、少了表）。
    """
    findings: list[_Finding] = []
    errors: list[_DoctorError] = []
    if not excel_path.exists():
        findings.append(("!", "工作簿", f"{excel_path} 不存在 —— 先跑 init 生成骨架"))
        return findings, errors

    try:
        workbook = load_workbook_file(excel_path)
    except CodeGenError as exc:
        # 打不开 = 环境问题（4）：不是配置写错了，而是这个文件现在读不了
        errors.append((_failure_code(exc), str(exc)))
        findings.append(("ERROR", "工作簿", str(exc).splitlines()[0]))
        return findings, errors

    # 先给默认值：下面任何一步失败都会跳进 except，而尾部还要用这三个
    recorded: dict[str, str] = {}
    now_input = ""
    now_output = ""
    try:
        check_required_sheets(workbook, project)
        findings.append(("OK", "工作表", "、".join(workbook.sheetnames)))
        global_values = read_global_values(workbook, project)
        cases = read_cases(workbook, project, global_values=global_values)
        axis_label = _case_axis_label(project)
        findings.append(("OK", axis_label, f"{len(cases)} 个：{', '.join(case.name for case in cases)}"))
        empty = [case.name for case in cases if not case.explicit_values]
        if empty:
            findings.append(
                (
                    "!",
                    f"空 {axis_label}",
                    f"{', '.join(empty)} 整{'列' if project.excel.local_direction == 'horizontal' else '行'}"
                    "都是空的 —— 所有变量都会回落 default，"
                    "可能悄悄落进某个 case_filter（指南 §9.2）",
                )
            )

        for label, checker in (("取值约束", check_value_constraints), ("asserts", check_asserts)):
            try:
                checker(project, global_values, cases)
                findings.append(("OK", label, "全部满足"))
            except CodeGenError as exc:
                errors.append((_failure_code(exc), str(exc)))
                findings.append(("ERROR", label, str(exc).splitlines()[0]))

        recorded = read_metadata(workbook, project)
        now_input = input_fingerprint(global_values, cases)
        try:
            now_output = output_fingerprint(render_all(project, excel_path).results)
        except CodeGenError as exc:
            # 参数本身有问题（约束 / asserts）时渲染不出来 —— 那已经在上面报过了
            errors.append((_failure_code(exc), str(exc)))
            now_output = ""
    except CodeGenError as exc:
        # 缺表 / 表头不对 / 变量名重复：工作簿与配置对不上（ExcelError -> 1）。
        # 关键是**别把 traceback 甩到用户脸上** —— 报告里给一行可读的结论。
        errors.append((_failure_code(exc), str(exc)))
        findings.append(("ERROR", "工作簿", str(exc).splitlines()[0]))
    finally:
        workbook.close()

    if not recorded:
        findings.append(("!", "渲染记录", "没有 —— 还没跑过 render --write-excel"))
    elif not now_output:
        findings.append(("!", "渲染记录", "参数不满足校验规则，渲染不出来（先修上面的 ERROR）"))
    else:
        same_input = recorded.get("参数指纹") == now_input
        same_output = recorded.get("输出指纹") == now_output
        if same_input and same_output:
            findings.append(("OK", "渲染记录", f"与当前一致（参数指纹 {now_input}）"))
        elif same_input:
            findings.append(("!", "渲染记录", "参数没变但输出指纹不同 —— 模板或代码改过，重跑一次"))
        else:
            findings.append(("!", "渲染记录", "参数改过了 —— 快照模板需要重跑（公式模板会自动重算）"))

    scripts = excel_path.parent / f"{excel_path.stem}_render.sh"
    findings.append(
        ("OK" if scripts.exists() else "!", "一键脚本", f"{scripts.name} {'在' if scripts.exists() else '不在'}")
    )
    return findings, errors


@app.command("doctor")
def doctor_command(
    config: Path = typer.Option(
        ..., "--config", "-c", exists=True, dir_okay=False, readable=True, help="YAML 配置文件路径"
    ),
    excel: Path | None = typer.Option(None, "--excel", "-x", help="顺带体检这个工作簿（默认取配置中的 excel.output）"),
) -> None:
    """体检环境 / 配置 / 工作簿，把常见坑一次说清。

    退出码：YAML 加载失败 = **2**（配置写错了）、工作簿打不开 = **4**（环境问题）、
    其余有 ERROR = **1**。同时存在多类问题时取最靠外的那一类（4 环境 > 2 配置 > 1 其他）
    —— 工作簿都读不了的时候，先解决那个才有意义。详见 docs/cli.md 的「退出码」一节。
    """
    findings: list[_Finding] = []
    errors: list[_DoctorError] = []

    findings.extend(_doctor_environment())

    try:
        project = _load_project(config)
    except CodeGenError as exc:
        findings.append(("ERROR", "YAML", str(exc).splitlines()[0]))
        _render_doctor(findings)
        raise typer.Exit(code=_failure_code(exc)) from exc

    excel_path = _resolve_excel_path(project, excel)
    config_findings, config_errors = _doctor_config(project, excel_path)
    workbook_findings, workbook_errors = _doctor_workbook(project, excel_path)
    findings.extend(config_findings)
    findings.extend(workbook_findings)
    errors.extend(config_errors)
    errors.extend(workbook_errors)

    for warning in project.load_warnings:
        findings.append(("!", "extends", warning))

    _render_doctor(findings)
    for _code, problem in errors:
        error_console.print(f"    {problem}")
    if errors:
        raise typer.Exit(code=max(code for code, _ in errors))
    notes = sum(1 for level, _, _ in findings if level == "!")
    console.print(
        f"[bold green]OK[/] 体检完成：{sum(1 for level, _, _ in findings if level == 'OK')} 项通过"
        + (f"，{notes} 项值得留意" if notes else "，没有需要留意的")
    )


@app.command("examples")
def examples_command(
    copy_to: Path | None = typer.Option(None, "--copy", "-o", help="把内置示例拷到这个目录（每个示例一个子目录）"),
    only: str | None = typer.Option(None, "--only", help="只处理某一个示例（basic / nastran / abs_fpi）"),
    force: bool = typer.Option(False, "--force", help="目标已存在且非空时覆盖"),
) -> None:
    """列出随包发布的内置示例；加 --copy DIR 就拷出来直接用。

    示例随 wheel / sdist 一起发布 —— 不需要克隆仓库。每个示例都带一本**已经填好样例参数**
    的工作簿：公式模式下打开就能改、改完自动重算。
    """
    if copy_to is None:
        console.print(
            "[bold]内置示例[/]（随包发布 —— 不用克隆仓库；"
            "[cyan]spreadsheet-codegen examples --copy ./examples[/] 拷出来用）\n"
        )
        for item in example_pack.EXAMPLES:
            console.print(f"  [bold cyan]{item.name}[/]  {item.title}")
            console.print(f"    [dim]{item.summary}[/]")
            console.print(f"    入口   [green]{item.entry}[/]")
            console.print(f"    工作簿 {item.workbook}\n")
        return

    try:
        selected = [example_pack.find(only)] if only else list(example_pack.EXAMPLES)
        written = example_pack.copy_examples(copy_to, only=only, force=force)
    except CodeGenError as exc:
        raise _fail(exc) from exc

    console.print(f"[bold green]OK[/] 已拷贝 {len(written)} 个示例到 [cyan]{copy_to}[/]")
    for item in selected:
        console.print(f"  {item.name:9s} {item.workbook}   [dim]打开就能改参数[/]")

    first = selected[0]
    console.print(
        "\n下一步（可选）：打开上面那本 .xlsx 填参数；"
        "要把代码导出成文件时跑\n"
        f"  [cyan]spreadsheet-codegen render -c {copy_to / first.entry} "
        f"-x {copy_to / first.workbook} --outdir out[/]"
    )


def _render_doctor(findings: Sequence[_Finding]) -> None:
    table = Table(title="体检报告", header_style="bold cyan")
    table.add_column("", width=6)
    table.add_column("项目", style="bold")
    table.add_column("结论")
    for level, name, detail in findings:
        mark = {"OK": "[green]OK[/]", "!": "[yellow]![/]"}.get(level, "[red]ERROR[/]")
        table.add_row(mark, name, detail)
    console.print(table)
