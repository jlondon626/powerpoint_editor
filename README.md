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

# One-step duplication: returns a SlideProxy bound to the new slide
slide = editor.duplicate_template_slide('RESERVE_WATERFALL')
slide.edit_textbox_html('AOM Text', '<b>Updated</b>')
slide.edit_embedded_workbook_for_chart('Waterfall chart', ['A','B'], [1,2])
slide.edit_chart_data_multi_series(
    'Multi-series chart',
    ['Q1', 'Q2', 'Q3', 'Q4'],
    {
        'Revenue': [10, 20, 30, 40],
        'Cost': [4, 8, 12, 16],
    },
)

# Alternatively obtain a proxy for an existing slide number
proxy = editor.get_slide(5)
proxy.edit_textbox('Title', 'Hello')

editor.replace_text_variables({'Quarter': 'Q1 2026'})

editor.drop_section('template_slides')
editor.save('output.pptx')
```

Examples

- Duplicate a template and edit it in one step:

```python
slide = editor.duplicate_template_slide('RESERVE_WATERFALL')
slide.edit_textbox_html('AOM Text', '<b>Generated</b>')
```

- Get a `SlideProxy` for an existing slide:

```python
slide = editor.get_slide(3)
slide.edit_table_cell('SummaryTable', 0, 1, 'Updated')
```

The `SlideProxy` exposes convenience methods delegating to the editor for common
per-slide operations: `edit_textbox`, `edit_textbox_html`, `edit_textbox_runs`,
`find_text_variables`, `replace_text_variables`, `remove_shape`,
`edit_table_cell`, `edit_table_range`, `edit_chart_data`,
`edit_chart_data_multi_series`, `edit_waterfall_data`,
`edit_embedded_workbook_for_chart`, and `refresh_chart`.

- Find and replace text variables:

```python
variables = editor.find_text_variables()
editor.replace_text_variables({'Quarter': 'Q1 2026', 'Region': 'EMEA'})

slide = editor.get_slide(3)
slide.replace_text_variables({'Quarter': 'Q2 2026'})
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

Publishing

A release workflow is included at `.github/workflows/publish.yml`.
When you create a GitHub release and publish it, the workflow will build
and upload the package to PyPI using the `PYPI_API_TOKEN` secret.

To configure publishing:

1. Create a PyPI API token at https://pypi.org/manage/account/token/
2. Add it to your repository secrets as `PYPI_API_TOKEN`
3. Create a release on GitHub

Notes
- The `refresh_chart()` function uses COM automation to refresh chart
  visuals inside PowerPoint; this only works on Windows with PowerPoint
  installed.
- The package manipulates the raw PPTX (zip) contents; always test on
  copies of presentations before running on production files.

API Reference (summary)

PowerPointEditor
- **`PowerPointEditor(input_pptx: str)`**: Load a `.pptx` into memory. Exposes `files` (part bytes) and `_has_changes`.
- **`save(output_pptx: Optional[str]) -> str`**: Write the modified package to disk; returns output path.
- **Slide/template handling**
  - `find_slide_by_shape_name(shape_name: str) -> int`: Find a slide containing a named shape (returns slide number).
  - `duplicate_slide(template_slide_number: int, before_section_name: Optional[str]) -> int`: Duplicate a slide; returns new slide number.
  - `duplicate_template_slide(template_name: str, before_section_name: str='template_slides') -> SlideProxy`: Duplicate a template slide marked with `TEMPLATE__<name>` and return a `SlideProxy` bound to the new slide for convenient edits.
  - `drop_section(section_name: str, delete_slide_parts: bool=True) -> list[int]`: Remove a section and its slides; by default also deletes now-unreferenced private slide dependencies.
  - `remove_shape_on_slide(slide_number: int, shape_name: str) -> None`: Remove a named shape or graphic frame from a slide.
- **Text editing**
  - `find_textbox_anywhere(textbox_name: str) -> dict`: Locate a textbox by name (returns `slide_number` and `slide_part`).
  - `edit_textbox_on_slide(slide_number: int, textbox_name: str, new_text: str) -> None`: Replace plain text (newlines -> paragraphs).
  - `edit_textbox_html_on_slide(slide_number: int, textbox_name: str, html: str) -> None`: Insert simple markup (`<b>` and `<br/>`).
  - `edit_textbox_runs_on_slide(slide_number: int, textbox_name: str, paragraphs: list[list[tuple[str,bool]]]) -> None`: Provide run-level `(text, bold)` paragraphs.
  - `find_text_variables() -> list[dict]`: Find `{Variable}` placeholders in slide text across the presentation.
  - `find_text_variables_on_slide(slide_number: int) -> list[dict]`: Find `{Variable}` placeholders on one slide.
  - `replace_text_variables(variables: dict[str, object]) -> int`: Replace `{Variable}` placeholders across the presentation; returns replacement count.
  - `replace_text_variables_on_slide(slide_number: int, variables: dict[str, object]) -> int`: Replace `{Variable}` placeholders on one slide.
  - Backward-compatible helpers: `edit_textbox(...)`, `edit_textbox_html(...)` operate anywhere.
- **Chart editing**
  - `find_chart_on_slide(slide_number: int, chart_name: str) -> dict`: Locate a chart frame and return `chart_part` and rel information.
  - `find_chart_anywhere(chart_name: str) -> dict`: Search all slides for a chart by name.
  - `edit_chart_data_on_slide(slide_number: int, chart_name: str, categories: list[str], values: list[float]) -> None`: Update chart caches and the embedded workbook for a regular chart when one is present.
  - `edit_chart_data_multi_series_on_slide(slide_number: int, chart_name: str, categories: list[str], series: dict[str, list[float]], sheet_name: Optional[str]=None) -> None`: Update a regular chart that already has multiple series; the template chart must contain the same number of series as the provided mapping.
  - `edit_embedded_workbook_for_chart_on_slide(slide_number: int, chart_name: str, categories: list[str], values: list[float], sheet_name: Optional[str]=None, subtotal_indices: Optional[list[int]]=None) -> None`: Update the embedded Excel file and chart XML (supports chartex subtotal indices).
  - Backward-compatible helpers: `edit_chart_data(...)`, `edit_chart_data_multi_series(...)`, `edit_waterfall_data(...)`, `edit_embedded_workbook_for_chart(...)`.
  - `refresh_chart(chart_name: str, output_pptx: str) -> None`: Attempt to refresh chart visuals using PowerPoint COM automation (Windows only).
- **Tables**
  - `find_table_on_slide(slide_number: int, table_name: str) -> dict`: Locate a table graphicFrame by shape name.
  - `find_table_anywhere(table_name: str) -> dict`: Search all slides for a table.
  - `edit_table_cell_on_slide(slide_number: int, table_name: str, row: int, col: int, new_text: str) -> None`: Edit a single cell (0-based indices).
  - `edit_table_range_on_slide(slide_number: int, table_name: str, data: list[list[str]]) -> None`: Write a 2D list of strings into the table (must fit table size).
- **Debug/list utilities**
  - `list_all_textboxes()`: Print named textboxes in the presentation.
  - `list_graphic_frames()`: Print graphic frames and chart/table presence per slide.
  - `list_sections()`: Print section names and slide counts.
  - `dump_chartex_debug(chart_name: str)`: Print raw chart XML and rels for debugging.

See `xmlppt/editor.py` for full docstrings and implementation details.
