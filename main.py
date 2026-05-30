"""Example entrypoint for the xmlppt package.

This file provides a small CLI to exercise the PowerPointEditor when
an input presentation is available. For programmatic usage prefer
importing `xmlppt.PowerPointEditor` from your code.
"""

from xmlppt import PowerPointEditor
import argparse
import os
import sys


def run_example(input_pptx: str, output_pptx: str | None = None) -> str:
    editor = PowerPointEditor(input_pptx=input_pptx)

    # Simple diagnostics
    editor.list_sections()
    editor.list_all_textboxes()
    editor.list_graphic_frames()

    # This example expects a template marker slide named
    # TEMPLATE__RESERVE_WATERFALL in a section called 'template_slides'.
    try:
        new_slide = editor.duplicate_template_slide(
            template_name="RESERVE_WATERFALL",
            before_section_name="template_slides",
        )

        editor.edit_textbox_html_on_slide(
            slide_number=new_slide,
            textbox_name="AOM Text",
            html=(
                "This is a generated slide for <b>Class A</b>.<br>"
                "The text has been edited after duplication."
            ),
        )

        editor.edit_embedded_workbook_for_chart_on_slide(
            slide_number=new_slide,
            chart_name="Waterfall chart",
            categories=[
                "2024 Q4",
                "AvE",
                "Change in\nPremium",
                "Specifics",
                "CATs",
                "2025 Q4",
            ],
            values=[1000, -120, 80, -14, -25, 921],
            subtotal_indices=[0, 5],
        )

    except Exception as exc:
        print(f"Example actions skipped due to error: {exc}")

    output = editor.save(output_pptx) if output_pptx else editor.save()
    return output


def main(argv: list[str] | None = None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    parser = argparse.ArgumentParser(description="xmlppt example CLI")
    parser.add_argument("--input", "-i", help="Input .pptx file to operate on")
    parser.add_argument("--output", "-o", help="Output .pptx file (optional)")
    parser.add_argument("--run-example", action="store_true", help="Run the bundled example operations")

    args = parser.parse_args(argv)

    if not args.input:
        parser.print_help()
        print("\nExample: python main.py --input example.pptx --run-example")
        return 1

    if not os.path.exists(args.input):
        print(f"Input file not found: {args.input}")
        return 2

    if args.run_example:
        out = run_example(args.input, args.output)
        print(f"Created: {out}")
    else:
        print("No action requested; use --run-example to run the sample flow.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
    def __init__(self, input_pptx: str):
        self.input_pptx = input_pptx
        self.files = self._load_pptx_files(input_pptx)
        self._has_changes = False

    @staticmethod
    def _normalize_relationship_target(base_part: str, target: str) -> str:
        base_dir = posixpath.dirname(base_part)
        return posixpath.normpath(posixpath.join(base_dir, target))

    def _load_pptx_files(self, pptx_path: str) -> dict[str, bytes]:
        with ZipFile(pptx_path, "r") as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    def _write_pptx_files(self, output_pptx: str) -> None:
        with ZipFile(output_pptx, "w") as archive:
            for name, data in self.files.items():
                archive.writestr(name, data)

    def _default_output_name(self, input_pptx: str) -> str:
        base, ext = os.path.splitext(input_pptx)
        return f"{base}_updated{ext}"

    def _find_textbox_shape_in_slide(self, slide_root, textbox_name: str):
        wanted = textbox_name.strip().casefold()
        for shape in slide_root.xpath(".//p:sp", namespaces=NS):
            cNvPr = shape.xpath("./p:nvSpPr/p:cNvPr", namespaces=NS)
            if not cNvPr:
                continue

            actual_name = (cNvPr[0].get("name") or "").strip().casefold()
            if actual_name == wanted:
                return shape
        return None

    def _append_text_run(self, paragraph, text: str, bold: bool = False) -> None:
        run = etree.SubElement(paragraph, f"{{{A_NS}}}r")
        if bold:
            etree.SubElement(run, f"{{{A_NS}}}rPr", b="1")

        text_element = etree.SubElement(run, f"{{{A_NS}}}t")
        if text.startswith(" ") or text.endswith(" "):
            text_element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_element.text = text

    def _replace_textbox_runs(self, txBody_elem, paragraphs: list[list[tuple[str, bool]]]) -> None:
        for paragraph in txBody_elem.xpath("./a:p", namespaces=NS):
            txBody_elem.remove(paragraph)

        for runs in paragraphs:
            paragraph = etree.SubElement(txBody_elem, f"{{{A_NS}}}p")
            if not runs:
                self._append_text_run(paragraph, "", False)
            else:
                for text, bold in runs:
                    self._append_text_run(paragraph, text, bold)

    def _text_to_paragraph_runs(self, text: str) -> list[list[tuple[str, bool]]]:
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        return [[(line, False)] for line in normalized.split("\n")]

    def _parse_textbox_markup(self, markup: str) -> list[list[tuple[str, bool]]]:
        normalized = markup.replace("\r\n", "\n").replace("\r", "\n")
        normalized = re.sub(r"(?i)<br\s*/?>", "\n", normalized)

        paragraphs: list[list[tuple[str, bool]]] = []
        for line in normalized.split("\n"):
            runs: list[tuple[str, bool]] = []
            last_index = 0
            for match in re.finditer(r"(?i)<b>(.*?)</b>", line, flags=re.DOTALL):
                if match.start() > last_index:
                    runs.append((line[last_index:match.start()], False))
                runs.append((match.group(1), True))
                last_index = match.end()

            if last_index < len(line):
                runs.append((line[last_index:], False))
            if not runs:
                runs.append(("", False))
            paragraphs.append(runs)

        return paragraphs

    def find_chart_anywhere(self, chart_name: str) -> dict:
        wanted = chart_name.strip().casefold()

        slide_parts = [
            (int(m.group(1)), name)
            for name in self.files
            if (m := SLIDE_RE.match(name))
        ]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            rels_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
            if rels_part not in self.files:
                continue

            slide_root = etree.fromstring(self.files[slide_part])
            rels_root = etree.fromstring(self.files[rels_part])
            frames = slide_root.xpath(".//p:graphicFrame", namespaces=NS)

            for frame in frames:
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

        raise ValueError(f"Chart named '{chart_name}' not found anywhere in presentation")

    def find_textbox_anywhere(self, textbox_name: str) -> dict:
        wanted = textbox_name.strip().casefold()

        slide_parts = [
            (int(m.group(1)), name)
            for name in self.files
            if (m := SLIDE_RE.match(name))
        ]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            slide_root = etree.fromstring(self.files[slide_part])
            shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)
            if shape is not None:
                return {
                    "slide_number": slide_number,
                    "slide_part": slide_part,
                }

        raise ValueError(f"Textbox named '{textbox_name}' not found anywhere in presentation")

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
            if rel.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package":
                return self._normalize_relationship_target(chart_part, rel.get("Target", ""))
        return None

    def _update_chartex_chart(
        self,
        chart_root,
        categories: list[str],
        values: list[float],
        subtotal_indices: list[int] | None = None,
    ) -> None:
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
        str_caches = chart_root.xpath("//c:cat/c:strCache", namespaces=NS)
        if str_caches:
            self._replace_str_cache(str_caches[0], categories)

        num_caches = chart_root.xpath("//c:val/c:numCache", namespaces=NS)
        if num_caches:
            self._replace_num_cache(num_caches[0], values)

    def _refresh_powerpoint_chart(self, output_pptx: str, chart_name: str) -> None:
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            presentation = ppt_app.Presentations.Open(os.path.abspath(output_pptx))
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.Name == chart_name and shape.HasChart:
                        shape.Chart.Refresh()
                        presentation.Save()
                        presentation.Close()
                        ppt_app.Quit()
                        return
            presentation.Close()
            ppt_app.Quit()
            print(f"Warning: Chart '{chart_name}' not found for refresh")
        except Exception as e:
            print(f"Warning: Could not refresh chart in PowerPoint: {e}")

    def edit_waterfall_data(
        self,
        chart_name: str,
        categories: list[str],
        values: list[float],
    ) -> None:
        """Update an older waterfall chart by editing its cached chart points."""
        if len(categories) != len(values):
            raise ValueError("categories and values must have the same length")

        found = self.find_chart_anywhere(chart_name)
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

        self.files[chart_part] = etree.tostring(
            chart_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )
        self._has_changes = True

    def edit_embedded_workbook_for_chart(
        self,
        chart_name: str,
        categories: list[str],
        values: list[float],
        sheet_name: str | None = None,
        subtotal_indices: list[int] | None = None,
    ) -> None:
        """Edit both the embedded workbook and the chart XML for a single chart."""
        if len(categories) != len(values):
            raise ValueError("categories and values must have the same length")

        found = self.find_chart_anywhere(chart_name)
        chart_part = found["chart_part"]
        rels_part = f"{posixpath.dirname(chart_part)}/_rels/{posixpath.basename(chart_part)}.rels"
        rels_root = etree.fromstring(self.files[rels_part])
        workbook_target = self._get_embedded_workbook_target(chart_part, rels_root)

        if not workbook_target:
            raise ValueError("No embedded workbook relationship found for this chart")
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

        chart_root = etree.fromstring(self.files[chart_part])
        self._update_chartex_chart(chart_root, categories, values, subtotal_indices)
        self._update_regular_chart_cache(chart_root, categories, values)
        self.files[chart_part] = etree.tostring(chart_root, encoding="utf-8", xml_declaration=True)
        self._has_changes = True

    def edit_textbox(
        self,
        textbox_name: str,
        new_text: str,
    ) -> None:
        """Edit the text content of a textbox by name."""
        found = self.find_textbox_anywhere(textbox_name)
        slide_part = found["slide_part"]
        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)
        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found in slide '{slide_part}'")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        txBody_elem = txBody[0]

        original_text = ""
        for p in txBody_elem.xpath("./a:p", namespaces=NS):
            for r in p.xpath("./a:r", namespaces=NS):
                for t in r.xpath("./a:t", namespaces=NS):
                    original_text += t.text or ""
        print(f"Original text in '{textbox_name}': '{original_text}'")

        paragraphs = self._text_to_paragraph_runs(new_text)
        self._replace_textbox_runs(txBody_elem, paragraphs)

        print(f"Updated text in '{textbox_name}' to: '{new_text}'")

        self.files[slide_part] = etree.tostring(
            slide_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )
        self._has_changes = True

    def edit_textbox_html(
        self,
        textbox_name: str,
        html: str,
    ) -> None:
        """Edit a textbox using simple HTML-like markup.

        Supported tags:
        - <b>bold</b>
        - <br> or <br/> for line breaks
        - \n for new paragraphs
        """
        found = self.find_textbox_anywhere(textbox_name)
        slide_part = found["slide_part"]
        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)
        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found in slide '{slide_part}'")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        paragraphs = self._parse_textbox_markup(html)
        self._replace_textbox_runs(txBody[0], paragraphs)

        self.files[slide_part] = etree.tostring(
            slide_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )
        self._has_changes = True

    def edit_textbox_runs(
        self,
        textbox_name: str,
        runs: list[tuple[str, bool]],
    ) -> None:
        """Edit a textbox using text runs with optional bold formatting."""
        found = self.find_textbox_anywhere(textbox_name)
        slide_part = found["slide_part"]
        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)
        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found in slide '{slide_part}'")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        self._replace_textbox_runs(txBody[0], runs)

        self.files[slide_part] = etree.tostring(
            slide_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )
        self._has_changes = True

    def list_all_textboxes(self) -> None:
        """List all textbox names found in the presentation."""
        slide_parts = [
            (int(m.group(1)), name)
            for name in self.files
            if (m := SLIDE_RE.match(name))
        ]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            slide_root = etree.fromstring(self.files[slide_part])
            shapes = slide_root.xpath(".//p:sp", namespaces=NS)

            for shape in shapes:
                cNvPr = shape.xpath("./p:nvSpPr/p:cNvPr", namespaces=NS)
                if not cNvPr:
                    continue
                name = cNvPr[0].get("name") or ""
                txBody = shape.xpath("./p:txBody", namespaces=NS)
                if txBody:
                    print(f"Slide {slide_number}: Textbox name='{name}'")

    def list_graphic_frames(self) -> None:
        slide_parts = [
            (int(m.group(1)), name)
            for name in self.files
            if (m := SLIDE_RE.match(name))
        ]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            slide_root = etree.fromstring(self.files[slide_part])
            frames = slide_root.xpath(".//p:graphicFrame", namespaces=NS)
            if not frames:
                continue

            print(f"\nSlide {slide_number}")
            for index, frame in enumerate(frames, start=1):
                cNvPr = frame.xpath("./p:nvGraphicFramePr/p:cNvPr", namespaces=NS)
                shape_name = cNvPr[0].get("name") if cNvPr else "<no name>"
                chart_refs = frame.xpath(".//c:chart", namespaces=NS)
                print(f"  {index}. name={shape_name!r}, is_chart={bool(chart_refs)}")

    def dump_chartex_debug(self, chart_name: str) -> None:
        found = self.find_chart_anywhere(chart_name)
        chart_part = found["chart_part"]
        chart_filename = posixpath.basename(chart_part)
        chart_dir = posixpath.dirname(chart_part)
        rels_part = f"{chart_dir}/_rels/{chart_filename}.rels"

        print("FOUND:", found)
        print()

        print("=== CHART PART ===")
        print(chart_part)
        print(self.files[chart_part].decode("utf-8", errors="ignore")[:8000])
        print()

        if rels_part in self.files:
            print("=== CHART RELS ===")
            print(rels_part)
            print(self.files[rels_part].decode("utf-8", errors="ignore")[:8000])
        else:
            print("No chart rels part found:", rels_part)

    def save(self, output_pptx: str) -> str:
        self._write_pptx_files(output_pptx)
        self._has_changes = False
        return output_pptx

    def refresh_chart(self, chart_name: str, output_pptx: str) -> None:
        self._refresh_powerpoint_chart(output_pptx, chart_name)


SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")

if __name__ == "__main__":
    editor = PowerPointEditor(input_pptx="example.pptx")

    # Example: edit a textbox in memory and save once at the end
    editor.edit_textbox_html(
        textbox_name="AOM Text",
        html="This is an <b>example</b> of editing a textbox with <b>bold</b> formatting and line breaks.<br>New line here.",
    )
    # editor.save("example_with_text.pptx")

    # Example: update the embedded chart workbook and chart XML
    editor.edit_embedded_workbook_for_chart(
        chart_name="Waterfall chart",
        categories=["2024 Q4", "AvE", "Change in\n Premium", "Specifics", "CATs", "2025 Q4"],
        values=[1000, -120, 80, -14, -25, 921],
        subtotal_indices=[0, 5],
    )
    editor.save("example_with_totals.pptx")

    print("Done")