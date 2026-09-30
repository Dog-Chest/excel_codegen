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

from .derived import (
    DerivedError,
    DerivedNotTranslatable,
    derived_variables,
    evaluate_derived,
    is_derived,
    is_translatable,
    ordered_derived,
)
from .derived import (
    validate_config as validate_derived,
)
from .excel_io import (
    check_required_sheets,
    create_template,
    input_fingerprint,
    load_workbook_file,
    output_fingerprint,
    read_cases,
    read_global_values,
    read_group_members,
    read_metadata,
    write_howto_sheet,
    write_results,
    write_run_scripts,
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
    check_asserts,
    collect_variables,
    compile_asserts,
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

__version__ = "0.9.0"

__all__ = [
    "FIRST_CASE_COLUMN",
    "CaseData",
    "CodeGenError",
    "ConfigError",
    "DerivedError",
    "DerivedNotTranslatable",
    "ExcelConfig",
    "ExcelError",
    "ExcelSheets",
    "FormulaError",
    "FormulaEvalError",
    "ProjectConfig",
    "RenderError",
    "RenderOutput",
    "RenderResult",
    "TemplateDef",
    "VarValue",
    "VariableDef",
    "VariablesConfig",
    "__version__",
    "build_context",
    "build_environment",
    "case_matches",
    "check_asserts",
    "check_required_sheets",
    "check_value_constraints",
    "collect_variables",
    "column_index_to_letter",
    "column_letter_to_index",
    "compile_asserts",
    "compile_case_filter",
    "compile_formulas",
    "compile_line",
    "create_template",
    "derived_variables",
    "evaluate_derived",
    "evaluate_formula",
    "evaluate_template_values",
    "excel_literal",
    "export_files",
    "fingerprint",
    "input_fingerprint",
    "is_derived",
    "is_translatable",
    "load_config",
    "load_workbook_file",
    "ordered_derived",
    "output_fingerprint",
    "parse_cell",
    "pvs",
    "read_cases",
    "read_global_values",
    "read_group_members",
    "read_metadata",
    "render_all",
    "render_template",
    "to_text",
    "validate_derived",
    "validate_template",
    "write_howto_sheet",
    "write_results",
    "write_run_scripts",
]
