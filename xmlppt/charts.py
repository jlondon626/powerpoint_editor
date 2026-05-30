from io import BytesIO
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
import posixpath
import re
import win32com.client
from .constants import *


class ChartMixin:
    def find_chart_on_slide(self, slide_number: int, chart_name: str) -> dict:
        """Locate a chart on a specific slide by shape name.

        Args:
            slide_number: 1-based slide index to search.
            chart_name: The shape name of the chart (case-insensitive).

        Returns:
            Dict containing `slide_number`, `slide_part`, `chart_part`, `rel_id`, and `chart_kind`.

        Raises:
            FileNotFoundError or ValueError when parts or chart are missing.
        """

        wanted = chart_name.strip().casefold()

        slide_part = f"ppt/slides/slide{slide_number}.xml"
        rels_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        if rels_part not in self.files:
            raise FileNotFoundError(f"Slide relationships not found: {rels_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        rels_root = etree.fromstring(self.files[rels_part])

        for frame in slide_root.xpath(".//p:graphicFrame", namespaces=NS):
            cNvPr = frame.xpath("./p:nvGraphicFramePr/p:cNvPr", namespaces=NS)
            if not cNvPr:
                continue

            actual_name = (cNvPr[0].get("name") or "").strip().casefold()
            if actual_name != wanted:
                continue

            chart_refs = frame.xpath(".//c:chart", namespaces=NS)
            chart_kind = "c"

            if not chart_refs:
                chart_refs = frame.xpath(".//cx:chart", namespaces=NS)
                chart_kind = "cx"

            if not chart_refs:
                continue

            rel_id = chart_refs[0].get(f"{{{R_NS}}}id")
            if not rel_id:
                continue

            rel = rels_root.xpath(f"./pr:Relationship[@Id='{rel_id}']", namespaces=NS)
            if not rel:
                continue

            target = rel[0].get("Target")
            if not target:
                continue

            return {
                "slide_number": slide_number,
                "slide_part": slide_part,
                "chart_part": self._normalize_relationship_target(slide_part, target),
                "rel_id": rel_id,
                "chart_kind": chart_kind,
            }

        raise ValueError(f"Chart named '{chart_name}' not found on slide {slide_number}")

    def find_chart_anywhere(self, chart_name: str) -> dict:
        """Find the first chart in the presentation with the given name.

        Args:
            chart_name: Chart shape name to search for (case-insensitive).

        Returns:
            The same dict returned by `find_chart_on_slide`.
        """

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, _slide_part in slide_parts:
            try:
                return self.find_chart_on_slide(slide_number, chart_name)
            except ValueError:
                continue

        raise ValueError(f"Chart named '{chart_name}' not found anywhere in presentation")

    def _replace_str_cache(self, str_cache, labels: list[str]) -> None:
        for pt in str_cache.xpath("./c:pt", namespaces=NS):
            str_cache.remove(pt)

        pt_count = str_cache.find(f"{{{C_NS}}}ptCount")
        if pt_count is None:
            pt_count = etree.Element(f"{{{C_NS}}}ptCount")
            str_cache.insert(0, pt_count)

        pt_count.set("val", str(len(labels)))

        for index, label in enumerate(labels):
            pt = etree.SubElement(str_cache, f"{{{C_NS}}}pt", idx=str(index))
            etree.SubElement(pt, f"{{{C_NS}}}v").text = str(label)

    def _replace_num_cache(self, num_cache, values: list[float]) -> None:
        for pt in num_cache.xpath("./c:pt", namespaces=NS):
            num_cache.remove(pt)

        pt_count = num_cache.find(f"{{{C_NS}}}ptCount")
        if pt_count is None:
            pt_count = etree.Element(f"{{{C_NS}}}ptCount")
            num_cache.insert(0, pt_count)

        pt_count.set("val", str(len(values)))

        for index, value in enumerate(values):
            pt = etree.SubElement(num_cache, f"{{{C_NS}}}pt", idx=str(index))
            etree.SubElement(pt, f"{{{C_NS}}}v").text = str(value)

    def _get_embedded_workbook_target(self, chart_part: str, rels_root) -> str | None:
        for rel in rels_root.xpath("./pr:Relationship", namespaces=NS):
            if rel.get("Type") == REL_TYPE_PACKAGE:
                return self._normalize_relationship_target(chart_part, rel.get("Target", ""))

        return None

    def _chart_rels_part(self, chart_part: str) -> str:
        return f"{posixpath.dirname(chart_part)}/_rels/{posixpath.basename(chart_part)}.rels"

    def _quote_sheet_name_for_formula(self, sheet_name: str) -> str:
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", sheet_name):
            return sheet_name

        escaped = sheet_name.replace("'", "''")
        return f"'{escaped}'"

    def _update_embedded_workbook(self, workbook_target: str, categories: list[str], values: list[float], sheet_name: str | None = None) -> str:
        if workbook_target not in self.files:
            raise FileNotFoundError(f"Embedded workbook not found: {workbook_target}")

        wb_bytes = BytesIO(self.files[workbook_target])
        wb = load_workbook(wb_bytes)
        sheet = wb[sheet_name] if sheet_name else wb.active

        start_row = 2
        end_row = max(sheet.max_row, len(categories) + 1)

        for row in range(start_row, end_row + 1):
            sheet[f"A{row}"] = None
            sheet[f"B{row}"] = None

        for row_index, (category, value) in enumerate(zip(categories, values), start=start_row):
            sheet[f"A{row_index}"] = category
            sheet[f"B{row_index}"] = value

        out_wb = BytesIO()
        wb.save(out_wb)
        self.files[workbook_target] = out_wb.getvalue()

        return sheet.title

    def _update_embedded_workbook_multi_series(self, workbook_target: str, categories: list[str], series: dict[str, list[float]], sheet_name: str | None = None) -> str:
        if workbook_target not in self.files:
            raise FileNotFoundError(f"Embedded workbook not found: {workbook_target}")

        wb_bytes = BytesIO(self.files[workbook_target])
        wb = load_workbook(wb_bytes)
        sheet = wb[sheet_name] if sheet_name else wb.active

        max_col = max(sheet.max_column, len(series) + 1)
        max_row = max(sheet.max_row, len(categories) + 1)

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                sheet.cell(row=row, column=col).value = None

        sheet["A1"] = "Category"
        for row_index, category in enumerate(categories, start=2):
            sheet.cell(row=row_index, column=1).value = category

        for series_index, (series_name, values) in enumerate(series.items(), start=2):
            sheet.cell(row=1, column=series_index).value = series_name
            for row_index, value in enumerate(values, start=2):
                sheet.cell(row=row_index, column=series_index).value = value

        out_wb = BytesIO()
        wb.save(out_wb)
        self.files[workbook_target] = out_wb.getvalue()

        return sheet.title

    def _update_regular_chart_formula_ranges(self, chart_root, sheet_name: str, point_count: int) -> None:
        quoted_sheet_name = self._quote_sheet_name_for_formula(sheet_name)
        category_range = f"{quoted_sheet_name}!$A$2:$A${point_count + 1}"
        value_range = f"{quoted_sheet_name}!$B$2:$B${point_count + 1}"

        for formula in chart_root.xpath(".//c:cat//c:strRef/c:f", namespaces=NS):
            formula.text = category_range

        for formula in chart_root.xpath(".//c:val//c:numRef/c:f", namespaces=NS):
            formula.text = value_range

    def _replace_series_title(self, series_node, series_name: str, sheet_name: str, column_letter: str) -> None:
        title_formula = f"{self._quote_sheet_name_for_formula(sheet_name)}!${column_letter}$1"

        formulas = series_node.xpath("./c:tx/c:strRef/c:f", namespaces=NS)
        for formula in formulas:
            formula.text = title_formula

        title_caches = series_node.xpath("./c:tx/c:strRef/c:strCache", namespaces=NS)
        for title_cache in title_caches:
            self._replace_str_cache(title_cache, [series_name])

        title_values = series_node.xpath("./c:tx/c:v", namespaces=NS)
        for title_value in title_values:
            title_value.text = series_name

    def _update_regular_chart_multi_series(self, chart_root, categories: list[str], series: dict[str, list[float]], sheet_name: str) -> None:
        series_nodes = chart_root.xpath(".//c:ser", namespaces=NS)

        if len(series_nodes) != len(series):
            raise ValueError(f"Chart has {len(series_nodes)} series but {len(series)} series were provided")

        quoted_sheet_name = self._quote_sheet_name_for_formula(sheet_name)
        category_range = f"{quoted_sheet_name}!$A$2:$A${len(categories) + 1}"

        for series_index, (series_node, (series_name, values)) in enumerate(zip(series_nodes, series.items()), start=2):
            column_letter = get_column_letter(series_index)
            value_range = f"{quoted_sheet_name}!${column_letter}$2:${column_letter}${len(values) + 1}"

            self._replace_series_title(series_node, series_name, sheet_name, column_letter)

            cat_caches = series_node.xpath(".//c:cat//c:strCache", namespaces=NS)
            for cat_cache in cat_caches:
                self._replace_str_cache(cat_cache, categories)

            cat_formulas = series_node.xpath(".//c:cat//c:strRef/c:f", namespaces=NS)
            for formula in cat_formulas:
                formula.text = category_range

            val_caches = series_node.xpath(".//c:val//c:numCache", namespaces=NS)
            for val_cache in val_caches:
                self._replace_num_cache(val_cache, values)

            val_formulas = series_node.xpath(".//c:val//c:numRef/c:f", namespaces=NS)
            for formula in val_formulas:
                formula.text = value_range

    def _update_chartex_chart(self, chart_root, categories: list[str], values: list[float], subtotal_indices: list[int] | None = None) -> None:
        cat_lvl = chart_root.xpath("//cx:strDim[@type='cat']/cx:lvl", namespaces=NS)
        if cat_lvl:
            lvl = cat_lvl[0]
            lvl.set("ptCount", str(len(categories)))

            for pt in lvl.xpath("./cx:pt", namespaces=NS):
                lvl.remove(pt)

            for index, category in enumerate(categories):
                pt = etree.SubElement(lvl, f"{{{CX_NS}}}pt", {"idx": str(index)})
                pt.text = str(category)

            f_elem = lvl.getparent().xpath("./cx:f", namespaces=NS)
            if f_elem:
                f_elem[0].text = f"Sheet1!$A$2:$A${len(categories) + 1}"

        val_lvl = chart_root.xpath("//cx:numDim[@type='val']/cx:lvl", namespaces=NS)
        if val_lvl:
            lvl = val_lvl[0]
            lvl.set("ptCount", str(len(values)))

            for pt in lvl.xpath("./cx:pt", namespaces=NS):
                lvl.remove(pt)

            for index, value in enumerate(values):
                pt = etree.SubElement(lvl, f"{{{CX_NS}}}pt", {"idx": str(index)})
                pt.text = str(value)

            f_elem = lvl.getparent().xpath("./cx:f", namespaces=NS)
            if f_elem:
                f_elem[0].text = f"Sheet1!$B$2:$B${len(values) + 1}"

        if subtotal_indices is not None:
            subtotal_node = chart_root.xpath("//cx:layoutPr/cx:subtotals", namespaces=NS)
            if subtotal_node:
                subtotals = subtotal_node[0]

                for idx_elem in subtotals.xpath("./cx:idx", namespaces=NS):
                    subtotals.remove(idx_elem)

                for idx in subtotal_indices:
                    etree.SubElement(subtotals, f"{{{CX_NS}}}idx", {"val": str(idx)})

    def _update_regular_chart_cache(self, chart_root, categories: list[str], values: list[float]) -> None:
        str_caches = chart_root.xpath("//c:cat//c:strCache", namespaces=NS)
        if str_caches:
            self._replace_str_cache(str_caches[0], categories)

        num_caches = chart_root.xpath("//c:val//c:numCache", namespaces=NS)
        if num_caches:
            self._replace_num_cache(num_caches[0], values)

    def edit_chart_data_on_slide(self, slide_number: int, chart_name: str, categories: list[str], values: list[float]) -> None:
        """Update the data caches for a regular chart on a slide.

        This updates the string and numeric caches embedded in the chart XML so
        the chart will display the supplied `categories` and `values`.

        Args:
            slide_number: 1-based slide index.
            chart_name: Chart shape name on the slide.
            categories: List of category labels (strings).
            values: List of numeric values (floats).

        Raises:
            ValueError: if lengths mismatch or expected caches are not found.
        """

        if len(categories) != len(values):
            raise ValueError("categories and values must have the same length")

        found = self.find_chart_on_slide(slide_number, chart_name)
        chart_part = found["chart_part"]

        chart_root = etree.fromstring(self.files[chart_part])

        str_cache_nodes = chart_root.xpath(".//c:cat//c:strCache", namespaces=NS)
        num_cache_nodes = chart_root.xpath(".//c:val//c:numCache", namespaces=NS)

        if not str_cache_nodes:
            raise ValueError("No category string cache found")

        if not num_cache_nodes:
            raise ValueError("No value numeric cache found")

        self._replace_str_cache(str_cache_nodes[0], categories)
        self._replace_num_cache(num_cache_nodes[0], values)

        rels_part = self._chart_rels_part(chart_part)
        if rels_part in self.files:
            rels_root = etree.fromstring(self.files[rels_part])
            workbook_target = self._get_embedded_workbook_target(chart_part, rels_root)
            if workbook_target:
                actual_sheet_name = self._update_embedded_workbook(
                    workbook_target=workbook_target,
                    categories=categories,
                    values=values,
                )
                self._update_regular_chart_formula_ranges(
                    chart_root=chart_root,
                    sheet_name=actual_sheet_name,
                    point_count=len(categories),
                )

        self.files[chart_part] = etree.tostring(chart_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        self._has_changes = True

    def edit_waterfall_data_on_slide(self, slide_number: int, chart_name: str, categories: list[str], values: list[float]) -> None:
        """Backward-compatible alias for `edit_chart_data_on_slide`."""
        return self.edit_chart_data_on_slide(slide_number=slide_number, chart_name=chart_name, categories=categories, values=values)

    def edit_chart_data_multi_series_on_slide(self, slide_number: int, chart_name: str, categories: list[str], series: dict[str, list[float]], sheet_name: str | None = None) -> None:
        """Update a regular chart with multiple value series.

        The chart template must already contain the same number of series as
        provided in `series`. This updates both chart XML caches and the
        embedded workbook so PowerPoint's Edit Data view stays in sync.
        """

        if not series:
            raise ValueError("series must contain at least one series")

        for series_name, values in series.items():
            if len(categories) != len(values):
                raise ValueError(f"Series '{series_name}' length does not match categories length")

        found = self.find_chart_on_slide(slide_number, chart_name)
        chart_part = found["chart_part"]

        rels_part = self._chart_rels_part(chart_part)

        if rels_part not in self.files:
            raise FileNotFoundError(f"Chart rels part not found: {rels_part}")

        rels_root = etree.fromstring(self.files[rels_part])
        workbook_target = self._get_embedded_workbook_target(chart_part, rels_root)

        if not workbook_target:
            raise ValueError("No embedded workbook relationship found for this chart")

        actual_sheet_name = self._update_embedded_workbook_multi_series(
            workbook_target=workbook_target,
            categories=categories,
            series=series,
            sheet_name=sheet_name,
        )

        chart_root = etree.fromstring(self.files[chart_part])
        self._update_regular_chart_multi_series(
            chart_root=chart_root,
            categories=categories,
            series=series,
            sheet_name=actual_sheet_name,
        )

        self.files[chart_part] = etree.tostring(chart_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        self._has_changes = True

    def edit_embedded_workbook_for_chart_on_slide(self, slide_number: int, chart_name: str, categories: list[str], values: list[float], sheet_name: str | None = None, subtotal_indices: list[int] | None = None) -> None:
        """Update the embedded Excel workbook for a chart and adjust chart XML.

        This writes `categories` and `values` into the embedded workbook (Sheet1
        or named sheet) and updates the chart caches so the chart will reflect
        the new workbook contents. If the chart is a 'chartex' type, subtotal
        indices may be applied.

        Args:
            slide_number: 1-based slide index.
            chart_name: Chart shape name on the slide.
            categories: List of category labels.
            values: List of numeric values.
            sheet_name: Optional sheet name to target in the workbook.
            subtotal_indices: Optional list of subtotal index integers (for chartex)

        Raises:
            ValueError: if lengths mismatch or embedded workbook/chart parts missing.
        """

        if len(categories) != len(values):
            raise ValueError("categories and values must have the same length")

        found = self.find_chart_on_slide(slide_number, chart_name)
        chart_part = found["chart_part"]

        rels_part = self._chart_rels_part(chart_part)

        if rels_part not in self.files:
            raise FileNotFoundError(f"Chart rels part not found: {rels_part}")

        rels_root = etree.fromstring(self.files[rels_part])
        workbook_target = self._get_embedded_workbook_target(chart_part, rels_root)

        if not workbook_target:
            raise ValueError("No embedded workbook relationship found for this chart")

        actual_sheet_name = self._update_embedded_workbook(
            workbook_target=workbook_target,
            categories=categories,
            values=values,
            sheet_name=sheet_name,
        )

        chart_root = etree.fromstring(self.files[chart_part])

        self._update_chartex_chart(chart_root=chart_root, categories=categories, values=values, subtotal_indices=subtotal_indices)

        self._update_regular_chart_cache(chart_root=chart_root, categories=categories, values=values)
        self._update_regular_chart_formula_ranges(chart_root=chart_root, sheet_name=actual_sheet_name, point_count=len(categories))

        self.files[chart_part] = etree.tostring(chart_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        self._has_changes = True

    # Backward-compatible anywhere methods
    def edit_chart_data(self, chart_name: str, categories: list[str], values: list[float]) -> None:
        """Find a chart by name anywhere and update its regular chart data caches."""

        found = self.find_chart_anywhere(chart_name)
        self.edit_chart_data_on_slide(slide_number=found["slide_number"], chart_name=chart_name, categories=categories, values=values)

    def edit_chart_data_multi_series(self, chart_name: str, categories: list[str], series: dict[str, list[float]], sheet_name: str | None = None) -> None:
        """Find a chart by name anywhere and update multiple regular chart series."""

        found = self.find_chart_anywhere(chart_name)
        self.edit_chart_data_multi_series_on_slide(slide_number=found["slide_number"], chart_name=chart_name, categories=categories, series=series, sheet_name=sheet_name)

    def edit_waterfall_data(self, chart_name: str, categories: list[str], values: list[float]) -> None:
        """Backward-compatible alias for `edit_chart_data`."""
        return self.edit_chart_data(chart_name=chart_name, categories=categories, values=values)

    def edit_embedded_workbook_for_chart(self, chart_name: str, categories: list[str], values: list[float], sheet_name: str | None = None, subtotal_indices: list[int] | None = None) -> None:
        """Find a chart by name anywhere and update its embedded workbook.

        Args are the same as `edit_embedded_workbook_for_chart_on_slide`.
        """

        found = self.find_chart_anywhere(chart_name)
        self.edit_embedded_workbook_for_chart_on_slide(slide_number=found["slide_number"], chart_name=chart_name, categories=categories, values=values, sheet_name=sheet_name, subtotal_indices=subtotal_indices)

    # PowerPoint COM refresh
    def _refresh_powerpoint_chart(self, output_pptx: str, chart_name: str) -> None:
        ppt_app = None
        presentation = None

        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            presentation = ppt_app.Presentations.Open(os.path.abspath(output_pptx))

            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.Name == chart_name and shape.HasChart:
                        shape.Chart.Refresh()

            presentation.Save()

        except Exception as exc:
            print(f"Warning: Could not refresh chart in PowerPoint: {exc}")

        finally:
            try:
                if presentation is not None:
                    presentation.Close()
            except Exception:
                pass

            try:
                if ppt_app is not None:
                    ppt_app.Quit()
            except Exception:
                pass

    def refresh_chart(self, chart_name: str, output_pptx: str) -> None:
        """Attempt to refresh a chart using PowerPoint COM automation.

        This opens the `output_pptx` in a PowerPoint COM instance and calls
        `.Refresh()` on the matching chart shape. This requires PowerPoint to
        be available on the host (Windows) and may be a no-op if COM automation
        fails.

        Args:
            chart_name: The shape name of the chart in PowerPoint.
            output_pptx: Path to the file to open in PowerPoint.
        """

        self._refresh_powerpoint_chart(output_pptx, chart_name)
