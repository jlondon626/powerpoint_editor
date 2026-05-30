# xmlppt

A small utility package for editing PowerPoint (.pptx) files by
manipulating the Open XML package directly. It supports duplicating
template slides, editing named textboxes, updating embedded Excel
workbooks for charts, and basic listing/debug utilities.

This repository contains a single package `xmlppt` with the main
implementation in `xmlppt/editor.py` and a minimal `main.py` example
entrypoint.

Requirements
- Python 3.10+
- lxml
- openpyxl
- pywin32 (only required for `refresh_chart()` on Windows)

Install

Create a virtual environment and install dependencies:

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1   # Windows PowerShell
pip install -r requirements.txt
```

Install the package locally for development:

```bash
pip install -e .
```

Quick usage (programmatic)

```python
from xmlppt import PowerPointEditor

editor = PowerPointEditor('input.pptx')
editor.duplicate_template_slide('RESERVE_WATERFALL')
editor.edit_textbox_html('AOM Text', '<b>Updated</b>')
editor.save('output.pptx')
```

CLI example

Run the example entrypoint which shows basic diagnostics and attempts
to run a sample flow (expects a marker template slide):

```bash
python main.py --input example.pptx --run-example
```

Run tests

```bash
pytest -q
```

CI

A GitHub Actions workflow is included at `.github/workflows/python-package.yml`
that installs the package and runs the test suite on Windows.

Notes
- The `refresh_chart()` function uses COM automation to refresh chart
  visuals inside PowerPoint; this only works on Windows with PowerPoint
  installed.
- The package manipulates the raw PPTX (zip) contents; always test on
  copies of presentations before running on production files.
