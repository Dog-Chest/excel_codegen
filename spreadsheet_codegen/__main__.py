"""允许 ``python -m spreadsheet_codegen`` 调用 CLI。"""

from __future__ import annotations

from .cli import app

if __name__ == "__main__":  # pragma: no cover
    app()
