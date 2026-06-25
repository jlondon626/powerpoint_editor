from lxml import etree
from .constants import *


class TextMixin:
    VARIABLE_PATTERN = re.compile(r"\{([^{}]+)\}")

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

    def _set_text_element_value(self, text_element, text: str) -> None:
        space_attr = "{http://www.w3.org/XML/1998/namespace}space"

        if text.startswith(" ") or text.endswith(" "):
            text_element.set(space_attr, "preserve")
        elif text_element.get(space_attr) == "preserve":
            text_element.attrib.pop(space_attr)

        text_element.text = text

    def _replace_paragraph_variables(self, paragraph, variables: dict[str, object]) -> int:
        text_elements = paragraph.xpath(".//a:t", namespaces=NS)
        if not text_elements:
            return 0

        original_text = "".join(text_element.text or "" for text_element in text_elements)
        replacement_count = 0

        def replace_match(match) -> str:
            nonlocal replacement_count
            variable_name = match.group(1).strip()

            if variable_name not in variables:
                return match.group(0)

            replacement_count += 1
            return str(variables[variable_name])

        updated_text = self.VARIABLE_PATTERN.sub(replace_match, original_text)

        if replacement_count == 0 or updated_text == original_text:
            return replacement_count

        self._set_text_element_value(text_elements[0], updated_text)

        for text_element in text_elements[1:]:
            self._set_text_element_value(text_element, "")

        return replacement_count

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

    def find_textbox_anywhere(self, textbox_name: str) -> dict:
        """Find the first textbox (shape) with the given name anywhere in the presentation.

        Args:
            textbox_name: The textbox `cNvPr/@name` to locate (case-insensitive).

        Returns:
            A dict containing `slide_number` and `slide_part` for the found textbox.

        Raises:
            ValueError: if no matching textbox is found.
        """

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
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

    def find_text_variables_on_slide(self, slide_number: int) -> list[dict]:
        """Find `{Variable}` placeholders in all text on a specific slide.

        Args:
            slide_number: 1-based slide index to search.

        Returns:
            A list of matches. Each match includes `variable`, `placeholder`,
            `slide_number`, `slide_part`, and the containing paragraph `text`.
        """

        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        results: list[dict] = []

        for paragraph in slide_root.xpath(".//a:p", namespaces=NS):
            text = "".join(text_element.text or "" for text_element in paragraph.xpath(".//a:t", namespaces=NS))

            for match in self.VARIABLE_PATTERN.finditer(text):
                results.append(
                    {
                        "variable": match.group(1).strip(),
                        "placeholder": match.group(0),
                        "slide_number": slide_number,
                        "slide_part": slide_part,
                        "text": text,
                    }
                )

        return results

    def find_text_variables(self) -> list[dict]:
        """Find all `{Variable}` placeholders in slide text across the presentation."""

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        results: list[dict] = []
        for slide_number, _slide_part in slide_parts:
            results.extend(self.find_text_variables_on_slide(slide_number))

        return results

    def replace_text_variables_on_slide(self, slide_number: int, variables: dict[str, object]) -> int:
        """Replace `{Variable}` placeholders in all text on a specific slide.

        Args:
            slide_number: 1-based slide index to update.
            variables: Mapping of variable names to replacement values, e.g.
                `{"Quarter": "Q1 2026"}` replaces `{Quarter}`.

        Returns:
            Number of placeholders replaced. Unknown placeholders are left as-is.
        """

        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        replacement_count = 0

        for paragraph in slide_root.xpath(".//a:p", namespaces=NS):
            replacement_count += self._replace_paragraph_variables(paragraph, variables)

        if replacement_count:
            self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            self._has_changes = True

        return replacement_count

    def replace_text_variables(self, variables: dict[str, object]) -> int:
        """Replace `{Variable}` placeholders in slide text across the presentation.

        Args:
            variables: Mapping of variable names to replacement values, e.g.
                `{"Quarter": "Q1 2026"}` replaces every `{Quarter}`.

        Returns:
            Total number of placeholders replaced. Unknown placeholders are left as-is.
        """

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        replacement_count = 0
        for slide_number, _slide_part in slide_parts:
            replacement_count += self.replace_text_variables_on_slide(slide_number, variables)

        return replacement_count

    def edit_textbox_on_slide(self, slide_number: int, textbox_name: str, new_text: str) -> None:
        """Replace the text of a textbox on a specific slide.

        Args:
            slide_number: 1-based slide index where the textbox lives.
            textbox_name: The textbox `cNvPr/@name` to edit (case-insensitive).
            new_text: Plain text to write; newline characters create new paragraphs.

        Raises:
            FileNotFoundError: if the slide part is missing.
            ValueError: if the named textbox or text body is not found.
        """

        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)

        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found on slide {slide_number}")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        paragraphs = self._text_to_paragraph_runs(new_text)
        self._replace_textbox_runs(txBody[0], paragraphs)

        self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        self._has_changes = True

    def edit_textbox_html_on_slide(self, slide_number: int, textbox_name: str, html: str) -> None:
        """Edit a textbox on a slide using a small subset of HTML-like markup.

        Supported tags: `<b>...</b>` for bold and `<br/>` for line breaks. Text
        is converted into presentation runs and paragraphs.

        Args:
            slide_number: 1-based slide index.
            textbox_name: The textbox name to edit.
            html: Markup string to parse and insert.
        """

        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)

        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found on slide {slide_number}")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        paragraphs = self._parse_textbox_markup(html)
        self._replace_textbox_runs(txBody[0], paragraphs)

        self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        self._has_changes = True

    def edit_textbox_runs_on_slide(self, slide_number: int, textbox_name: str, paragraphs: list[list[tuple[str, bool]]]) -> None:
        """Edit a textbox by supplying explicit run structures.

        Args:
            slide_number: 1-based slide index.
            textbox_name: The textbox name to edit.
            paragraphs: A list of paragraphs, each paragraph is a list of
                `(text, bold)` tuples where `bold` is a bool.
        """

        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        shape = self._find_textbox_shape_in_slide(slide_root, textbox_name)

        if shape is None:
            raise ValueError(f"Textbox named '{textbox_name}' not found on slide {slide_number}")

        txBody = shape.xpath("./p:txBody", namespaces=NS)
        if not txBody:
            raise ValueError(f"No text body found in textbox '{textbox_name}'")

        self._replace_textbox_runs(txBody[0], paragraphs)

        self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        self._has_changes = True

    # Backward-compatible anywhere methods
    def edit_textbox(self, textbox_name: str, new_text: str) -> None:
        """Backward-compatible helper: find and edit the first matching textbox.

        Args:
            textbox_name: The textbox name to locate and edit.
            new_text: Plain text to insert.
        """

        found = self.find_textbox_anywhere(textbox_name)
        self.edit_textbox_on_slide(found["slide_number"], textbox_name, new_text)

    def edit_textbox_html(self, textbox_name: str, html: str) -> None:
        """Backward-compatible helper: find a textbox and edit it with HTML markup.

        Args:
            textbox_name: The textbox name to locate.
            html: Markup string to insert into the textbox.
        """

        found = self.find_textbox_anywhere(textbox_name)
        self.edit_textbox_html_on_slide(found["slide_number"], textbox_name, html)
