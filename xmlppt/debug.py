from lxml import etree
import posixpath
from .constants import *


class DebugMixin:
    def list_all_textboxes(self) -> None:
        """Print all named textboxes found in the presentation.

        This helper iterates slides and prints `cNvPr/@name` for shapes that
        contain a text body. Intended as a debugging utility.
        """

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            slide_root = etree.fromstring(self.files[slide_part])

            for shape in slide_root.xpath(".//p:sp", namespaces=NS):
                cNvPr = shape.xpath("./p:nvSpPr/p:cNvPr", namespaces=NS)
                if not cNvPr:
                    continue

                name = cNvPr[0].get("name") or ""
                txBody = shape.xpath("./p:txBody", namespaces=NS)

                if txBody:
                    print(f"Slide {slide_number}: Textbox name={name!r}")

    def list_graphic_frames(self) -> None:
        """Print graphic frames (charts/tables) found on each slide for debugging."""

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
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
                chartex_refs = frame.xpath(".//cx:chart", namespaces=NS)

                print(
                    f"  {index}. name={shape_name!r}, is_regular_chart={bool(chart_refs)}, is_chartex_chart={bool(chartex_refs)}"
                )

    def list_sections(self) -> None:
        """Print section names and slide counts from the presentation manifest."""

        presentation_root = etree.fromstring(self.files["ppt/presentation.xml"])
        sections = presentation_root.xpath(".//p14:section", namespaces=NS)

        if not sections:
            print("No sections found")
            return

        for index, section in enumerate(sections, start=1):
            name = section.get("name") or ""
            slide_ids = section.xpath("./p14:sldIdLst/p14:sldId", namespaces=NS)
            print(f"{index}. {name!r}: {len(slide_ids)} slides")

    def dump_chartex_debug(self, chart_name: str) -> None:
        """Debug helper: print raw chartex chart XML and rels for a chart.

        Args:
            chart_name: Name of the chart to dump.
        """

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

    # ============================================================
    # Table lookup and editing
