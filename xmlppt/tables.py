from lxml import etree
from .constants import *


class TableMixin:
    def find_table_on_slide(self, slide_number: int, table_name: str) -> dict:
        """Locate a table (graphicFrame containing an a:tbl) by its shape name on a slide.

        Args:
            slide_number: 1-based slide index to search.
            table_name: The shape name of the table (case-insensitive).

        Returns:
            Dict with `slide_number` and `slide_part` for the found table.

        Raises:
            FileNotFoundError: if the slide part is missing.
            ValueError: if the named table is not found.
        """
        wanted = table_name.strip().casefold()

        slide_part = f"ppt/slides/slide{slide_number}.xml"
        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])

        for frame in slide_root.xpath('.//p:graphicFrame', namespaces=NS):
            cNvPr = frame.xpath('./p:nvGraphicFramePr/p:cNvPr', namespaces=NS)
            if not cNvPr:
                continue

            actual_name = (cNvPr[0].get('name') or '').strip().casefold()
            if actual_name != wanted:
                continue

            # check if this graphicFrame contains a table
            tbl = frame.xpath('.//a:tbl', namespaces=NS)
            if not tbl:
                continue

            return {
                'slide_number': slide_number,
                'slide_part': slide_part,
            }

        raise ValueError(f"Table named '{table_name}' not found on slide {slide_number}")

    def find_table_anywhere(self, table_name: str) -> dict:
        """Find the first table with the given name anywhere in the presentation.

        Args:
            table_name: Table shape name to locate.

        Returns:
            Dict returned by `find_table_on_slide`.
        """

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, _slide_part in slide_parts:
            try:
                return self.find_table_on_slide(slide_number, table_name)
            except ValueError:
                continue

        raise ValueError(f"Table named '{table_name}' not found anywhere in presentation")

    def _set_table_cell_text(self, txBody_elem, text: str) -> None:
        """Helper: replace the text content of a table cell's txBody."""
        paragraphs = self._text_to_paragraph_runs(text)
        self._replace_textbox_runs(txBody_elem, paragraphs)

    def edit_table_cell_on_slide(self, slide_number: int, table_name: str, row: int, col: int, new_text: str) -> None:
        """Edit a single table cell by 0-based `row` and `col` indices on a slide.

        Args:
            slide_number: 1-based slide index containing the table.
            table_name: Table shape name to locate.
            row: Zero-based row index within the table.
            col: Zero-based column index within the row.
            new_text: Text to write into the cell.

        Raises:
            IndexError: if row/col are out of range.
            ValueError/FileNotFoundError when table or slide parts are missing.
        """
        info = self.find_table_on_slide(slide_number, table_name)
        slide_part = info['slide_part']

        slide_root = etree.fromstring(self.files[slide_part])

        # find the specific frame again and the table element
        for frame in slide_root.xpath('.//p:graphicFrame', namespaces=NS):
            cNvPr = frame.xpath('./p:nvGraphicFramePr/p:cNvPr', namespaces=NS)
            if not cNvPr:
                continue

            actual_name = (cNvPr[0].get('name') or '').strip().casefold()
            if actual_name != table_name.strip().casefold():
                continue

            tbls = frame.xpath('.//a:tbl', namespaces=NS)
            if not tbls:
                continue

            tbl = tbls[0]
            rows = tbl.xpath('./a:tr', namespaces=NS)
            if row < 0 or row >= len(rows):
                raise IndexError('row index out of range')

            cells = rows[row].xpath('./a:tc', namespaces=NS)
            if col < 0 or col >= len(cells):
                raise IndexError('col index out of range')

            cell = cells[col]
            txBody = cell.xpath('.//a:txBody', namespaces=NS)
            if not txBody:
                raise ValueError('Table cell has no text body')

            self._set_table_cell_text(txBody[0], new_text)

            # persist changes
            self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding='UTF-8', standalone='yes')
            self._has_changes = True
            return

        raise ValueError(f"Table named '{table_name}' not found on slide {slide_number}")

    def edit_table_range_on_slide(self, slide_number: int, table_name: str, data: list[list[str]]) -> None:
        """Write a 2D list of strings into a table on a slide.

        Args:
            slide_number: 1-based slide index containing the table.
            table_name: Table shape name to locate.
            data: 2D list of strings where each inner list is a table row.

        Raises:
            ValueError: if provided data exceeds table dimensions.
        """
        info = self.find_table_on_slide(slide_number, table_name)
        slide_part = info['slide_part']

        slide_root = etree.fromstring(self.files[slide_part])

        for frame in slide_root.xpath('.//p:graphicFrame', namespaces=NS):
            cNvPr = frame.xpath('./p:nvGraphicFramePr/p:cNvPr', namespaces=NS)
            if not cNvPr:
                continue

            actual_name = (cNvPr[0].get('name') or '').strip().casefold()
            if actual_name != table_name.strip().casefold():
                continue

            tbls = frame.xpath('.//a:tbl', namespaces=NS)
            if not tbls:
                continue

            tbl = tbls[0]
            rows = tbl.xpath('./a:tr', namespaces=NS)

            if len(data) > len(rows):
                raise ValueError('Provided data has more rows than the table')

            for r_idx, row_vals in enumerate(data):
                cells = rows[r_idx].xpath('./a:tc', namespaces=NS)
                if len(row_vals) > len(cells):
                    raise ValueError(f'Row {r_idx} has more columns than the table')

                for c_idx, val in enumerate(row_vals):
                    cell = cells[c_idx]
                    txBody = cell.xpath('.//a:txBody', namespaces=NS)
                    if not txBody:
                        raise ValueError('Table cell has no text body')
                    self._set_table_cell_text(txBody[0], val)

            # persist changes
            self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding='UTF-8', standalone='yes')
            self._has_changes = True
            return

        raise ValueError(f"Table named '{table_name}' not found on slide {slide_number}")
