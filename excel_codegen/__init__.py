"""excel_codegen —— Excel 模板参数填写 + Jinja2 代码生成器。

典型用法::

    from excel_codegen import load_config, create_template, render_all, write_results

    config = load_config("examples/example.yaml")
    create_template(config, "template.xlsx", cases=2, overwrite=True)
    # ... 用户在 Excel 中填写参数 ...
    output = render_all(config, "template.xlsx")
    write_results("template.xlsx", config, output.results)

命令行入口见 :mod:`excel_codegen.cli`（命令 ``excel-codegen``：init / render / validate / check）。
"""

from __future__ import annotations

from .excel_io import (
    check_required_sheets,
    create_template,
    input_fingerprint,
    load_workbook_file,
    output_fingerprint,
    read_cases,
    read_global_values,
    read_metadata,
    write_howto_sheet,
    write_results,
)
from .derived import (
    DerivedError,
    DerivedNotTranslatable,
    derived_variables,
    evaluate_derived,
    is_derived,
    is_translatable,
    ordered_derived,
    validate_config as validate_derived,
)
from .formula import (
    FormulaError,
    compile_formulas,
    compile_line,
    excel_literal,
)
from .formula_eval import (
    FormulaEvalError,
    evaluate_formula,
    evaluate_template_values,
)
from .models import (
    FIRST_CASE_COLUMN,
    CaseData,
    ExcelConfig,
    ExcelSheets,
    ProjectConfig,
    RenderResult,
    TemplateDef,
    VariableDef,
    VariablesConfig,
    load_config,
)
from .renderer import (
    RenderOutput,
    build_context,
    build_environment,
    case_matches,
    collect_variables,
    compile_case_filter,
    export_files,
    pvs,
    render_all,
    render_template,
    validate_template,
)
from .utils import (
    CodeGenError,
    ConfigError,
    ExcelError,
    RenderError,
    VarValue,
    column_index_to_letter,
    column_letter_to_index,
    fingerprint,
    parse_cell,
    to_text,
)

__version__ = "0.5.2"

__all__ = [
    "__version__",
    # derived（派生参数）
    "DerivedError",
    "DerivedNotTranslatable",
    "derived_variables",
    "evaluate_derived",
    "is_derived",
    "is_translatable",
    "ordered_derived",
    "validate_derived",
    # formula
    "FormulaError",
    "compile_formulas",
    "compile_line",
    "excel_literal",
    # formula eval
    "FormulaEvalError",
    "evaluate_formula",
    "evaluate_template_values",
    # models
    "FIRST_CASE_COLUMN",
    "CaseData",
    "ExcelConfig",
    "ExcelSheets",
    "ProjectConfig",
    "RenderResult",
    "TemplateDef",
    "VariableDef",
    "VariablesConfig",
    "load_config",
    # excel io
    "check_required_sheets",
    "create_template",
    "input_fingerprint",
    "load_workbook_file",
    "output_fingerprint",
    "read_cases",
    "read_global_values",
    "read_metadata",
    "write_howto_sheet",
    "write_results",
    # renderer
    "RenderOutput",
    "build_context",
    "build_environment",
    "case_matches",
    "collect_variables",
    "compile_case_filter",
    "export_files",
    "pvs",
    "render_all",
    "render_template",
    "validate_template",
    # utils
    "CodeGenError",
    "ConfigError",
    "ExcelError",
    "RenderError",
    "VarValue",
    "column_index_to_letter",
    "column_letter_to_index",
    "fingerprint",
    "parse_cell",
    "to_text",
]
