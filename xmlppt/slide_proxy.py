class SlideProxy:
    """A small convenience wrapper bound to a single slide number.

    Use `editor.get_slide(n)` to obtain an instance. Methods on the proxy
    delegate to the corresponding `PowerPointEditor` methods with the
    bound slide number.
    """

    def __init__(self, editor: "PowerPointEditor", slide_number: int):
        self._editor = editor
        self.slide_number = slide_number

    def edit_textbox(self, textbox_name: str, new_text: str) -> None:
        return self._editor.edit_textbox_on_slide(self.slide_number, textbox_name, new_text)

    def edit_textbox_html(self, textbox_name: str, html: str) -> None:
        return self._editor.edit_textbox_html_on_slide(self.slide_number, textbox_name, html)

    def edit_textbox_runs(self, textbox_name: str, paragraphs: list[list[tuple[str, bool]]]) -> None:
        return self._editor.edit_textbox_runs_on_slide(self.slide_number, textbox_name, paragraphs)

    def find_text_variables(self) -> list[dict]:
        return self._editor.find_text_variables_on_slide(self.slide_number)

    def replace_text_variables(self, variables: dict[str, object]) -> int:
        return self._editor.replace_text_variables_on_slide(self.slide_number, variables)

    def remove_shape(self, shape_name: str) -> None:
        return self._editor.remove_shape_on_slide(self.slide_number, shape_name)

    def edit_table_cell(self, table_name: str, row: int, col: int, new_text: str) -> None:
        return self._editor.edit_table_cell_on_slide(self.slide_number, table_name, row, col, new_text)

    def edit_table_range(self, table_name: str, data: list[list[str]]) -> None:
        return self._editor.edit_table_range_on_slide(self.slide_number, table_name, data)

    def __repr__(self) -> str:
        return f"<SlideProxy slide_number={self.slide_number}>"
    
    # Chart-related helpers bound to this slide
    def edit_chart_data(self, chart_name: str, categories: list[str], values: list[float]) -> None:
        return self._editor.edit_chart_data_on_slide(self.slide_number, chart_name, categories, values)

    def edit_chart_data_multi_series(self, chart_name: str, categories: list[str], series: dict[str, list[float]], sheet_name: str | None = None) -> None:
        return self._editor.edit_chart_data_multi_series_on_slide(self.slide_number, chart_name, categories, series, sheet_name=sheet_name)

    def edit_waterfall_data(self, chart_name: str, categories: list[str], values: list[float]) -> None:
        return self.edit_chart_data(chart_name, categories, values)

    def edit_embedded_workbook_for_chart(self, chart_name: str, categories: list[str], values: list[float], sheet_name: str | None = None, subtotal_indices: list[int] | None = None) -> None:
        return self._editor.edit_embedded_workbook_for_chart_on_slide(self.slide_number, chart_name, categories, values, sheet_name=sheet_name, subtotal_indices=subtotal_indices)

    def refresh_chart(self, chart_name: str, output_pptx: str) -> None:
        return self._editor.refresh_chart(chart_name, output_pptx)
