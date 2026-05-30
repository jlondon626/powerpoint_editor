from lxml import etree
import posixpath
from .constants import *
from .slide_proxy import SlideProxy


class SlideSectionMixin:
    def _find_insert_index_before_section(self, presentation_root, section_name: str) -> int | None:
        wanted = section_name.strip().casefold()

        sections = presentation_root.xpath(".//p14:section", namespaces=NS)
        if not sections:
            raise ValueError("No PowerPoint sections found in presentation.xml. Create a section named 'template_slides' in PowerPoint first.")

        target_section = None
        for section in sections:
            actual = (section.get("name") or "").strip().casefold()
            if actual == wanted:
                target_section = section
                break

        if target_section is None:
            raise ValueError(f"Section named '{section_name}' not found")

        section_slide_ids = target_section.xpath("./p14:sldIdLst/p14:sldId", namespaces=NS)
        if not section_slide_ids:
            raise ValueError(f"Section '{section_name}' contains no slides")

        first_template_slide_id = section_slide_ids[0].get("id")
        if not first_template_slide_id:
            raise ValueError(f"First slide in section '{section_name}' has no slide id")

        sld_id_lst = presentation_root.xpath("./p:sldIdLst", namespaces=NS)
        if not sld_id_lst:
            raise ValueError("No p:sldIdLst found in presentation.xml")

        all_slide_ids = sld_id_lst[0].xpath("./p:sldId", namespaces=NS)

        for index, sld_id in enumerate(all_slide_ids):
            if sld_id.get("id") == first_template_slide_id:
                return index

        raise ValueError(f"Could not locate first slide of section '{section_name}' in p:sldIdLst")

    def _add_slide_id_to_section_before(self, presentation_root, new_slide_id: str, before_section_name: str) -> None:
        wanted = before_section_name.strip().casefold()

        sections = presentation_root.xpath(".//p14:section", namespaces=NS)
        if not sections:
            return

        target_index = None
        for index, section in enumerate(sections):
            actual = (section.get("name") or "").strip().casefold()
            if actual == wanted:
                target_index = index
                break

        if target_index is None:
            raise ValueError(f"Section named '{before_section_name}' not found")

        if target_index == 0:
            raise ValueError("Cannot insert before section because it is the first section.")

        previous_section = sections[target_index - 1]

        sld_id_lst = previous_section.xpath("./p14:sldIdLst", namespaces=NS)
        if sld_id_lst:
            section_sld_id_lst = sld_id_lst[0]
        else:
            section_sld_id_lst = etree.SubElement(previous_section, f"{{{P14_NS}}}sldIdLst")

        new_section_sld_id = etree.SubElement(section_sld_id_lst, f"{{{P14_NS}}}sldId")
        new_section_sld_id.set("id", new_slide_id)

    def _add_slide_to_presentation(self, slide_part: str, before_section_name: str | None = None) -> int:
        presentation_part = "ppt/presentation.xml"
        presentation_rels_part = "ppt/_rels/presentation.xml.rels"

        presentation_root = etree.fromstring(self.files[presentation_part])
        presentation_rels_root = etree.fromstring(self.files[presentation_rels_part])

        new_rid = self._next_relationship_id(presentation_rels_root)

        rel = etree.SubElement(presentation_rels_root, f"{{{PR_NS}}}Relationship")
        rel.set("Id", new_rid)
        rel.set("Type", REL_TYPE_SLIDE)
        rel.set("Target", posixpath.relpath(slide_part, "ppt"))

        sld_id_lst = presentation_root.xpath("./p:sldIdLst", namespaces=NS)
        if not sld_id_lst:
            sld_id_lst_elem = etree.SubElement(presentation_root, f"{{{P_NS}}}sldIdLst")
        else:
            sld_id_lst_elem = sld_id_lst[0]

        new_slide_id = self._next_slide_id(presentation_root)

        new_sld_id = etree.Element(f"{{{P_NS}}}sldId")
        new_sld_id.set("id", new_slide_id)
        new_sld_id.set(f"{{{R_NS}}}id", new_rid)

        insert_index = None

        if before_section_name:
            insert_index = self._find_insert_index_before_section(presentation_root=presentation_root, section_name=before_section_name)

        if insert_index is None:
            sld_id_lst_elem.append(new_sld_id)
        else:
            sld_id_lst_elem.insert(insert_index, new_sld_id)

        if before_section_name:
            self._add_slide_id_to_section_before(presentation_root=presentation_root, new_slide_id=new_slide_id, before_section_name=before_section_name)

        self.files[presentation_part] = etree.tostring(presentation_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        self.files[presentation_rels_part] = etree.tostring(presentation_rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        slide_num_match = SLIDE_RE.match(slide_part)
        if not slide_num_match:
            raise ValueError(f"Unexpected slide part name: {slide_part}")

        return int(slide_num_match.group(1))

    # Deep-copy slide dependencies
    def _copy_related_part_if_needed(self, source_part: str, dest_part: str, rel_type: str) -> str | None:
        if rel_type in {REL_TYPE_SLIDE_LAYOUT, REL_TYPE_IMAGE}:
            return None

        if rel_type == REL_TYPE_CHART:
            new_part = self._next_numbered_part_name("ppt/charts", "chart", ".xml")
        elif rel_type == REL_TYPE_PACKAGE:
            new_part = self._next_numbered_part_name("ppt/embeddings", "Microsoft_Excel_Worksheet", ".xlsx")
        elif rel_type == REL_TYPE_CHART_STYLE:
            new_part = self._next_numbered_part_name("ppt/charts", "style", ".xml")
        elif rel_type == REL_TYPE_CHART_COLOR_STYLE:
            new_part = self._next_numbered_part_name("ppt/charts", "colors", ".xml")
        else:
            return None

        self._deep_copy_part(source_part, new_part)

        # Ensure copied parts are registered in [Content_Types].xml so PowerPoint
        # does not mark the package as corrupt when new chart/embed parts are added.
        if rel_type == REL_TYPE_CHART:
            chart_content_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
            self._add_content_type_override(new_part, chart_content_type)
        elif rel_type == REL_TYPE_PACKAGE:
            # Embedded workbooks generally use the default .xlsx mapping, but
            # adding an explicit override is harmless and keeps the package
            # consistent when new embedding parts are created.
            pkg_content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            self._add_content_type_override(new_part, pkg_content_type)
        elif rel_type == REL_TYPE_CHART_STYLE:
            style_content_type = "application/vnd.ms-office.chartstyle+xml"
            self._add_content_type_override(new_part, style_content_type)
        elif rel_type == REL_TYPE_CHART_COLOR_STYLE:
            color_style_content_type = "application/vnd.ms-office.chartcolorstyle+xml"
            self._add_content_type_override(new_part, color_style_content_type)

        dest_folder = posixpath.dirname(dest_part)
        return posixpath.relpath(new_part, dest_folder)

    def _deep_copy_part(self, source_part: str, dest_part: str) -> None:
        if source_part not in self.files:
            raise FileNotFoundError(f"Cannot copy missing part: {source_part}")

        self.files[dest_part] = self.files[source_part]

        source_rels_part = self._rels_part_for(source_part)
        dest_rels_part = self._rels_part_for(dest_part)

        if source_rels_part not in self.files:
            return

        rels_root = etree.fromstring(self.files[source_rels_part])

        for rel in rels_root.xpath("./pr:Relationship", namespaces=NS):
            target_mode = rel.get("TargetMode")
            if target_mode == "External":
                continue

            target = rel.get("Target")
            rel_type = rel.get("Type")

            if not target or not rel_type:
                continue

            source_related_part = self._normalize_relationship_target(source_part, target)

            if source_related_part not in self.files:
                continue

            new_target = self._copy_related_part_if_needed(source_part=source_related_part, dest_part=dest_part, rel_type=rel_type)

            if new_target:
                rel.set("Target", new_target)

        self.files[dest_rels_part] = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    # Template slide lookup and duplication
    def find_slide_by_shape_name(self, shape_name: str) -> int:
        """Locate the first slide that contains a shape with `name` equal to `shape_name`.

        Args:
            shape_name: The shape `cNvPr/@name` to search for (case-insensitive).

        Returns:
            The slide number (1-based) containing a matching shape.

        Raises:
            ValueError: if no slide contains a shape with the given name.
        """

        wanted = shape_name.strip().casefold()

        slide_parts = [(int(match.group(1)), name) for name in self.files if (match := SLIDE_RE.match(name))]
        slide_parts.sort(key=lambda item: item[0])

        for slide_number, slide_part in slide_parts:
            slide_root = etree.fromstring(self.files[slide_part])

            for shape in slide_root.xpath(".//p:sp", namespaces=NS):
                cNvPr = shape.xpath("./p:nvSpPr/p:cNvPr", namespaces=NS)
                if not cNvPr:
                    continue

                actual = (cNvPr[0].get("name") or "").strip().casefold()
                if actual == wanted:
                    return slide_number

            for frame in slide_root.xpath(".//p:graphicFrame", namespaces=NS):
                cNvPr = frame.xpath("./p:nvGraphicFramePr/p:cNvPr", namespaces=NS)
                if not cNvPr:
                    continue

                actual = (cNvPr[0].get("name") or "").strip().casefold()
                if actual == wanted:
                    return slide_number

        raise ValueError(f"Slide marker shape named '{shape_name}' not found")

    def remove_shape_on_slide(self, slide_number: int, shape_name: str) -> None:
        """Remove the first shape (or graphicFrame) with the given name on a slide.

        Args:
            slide_number: 1-based slide index.
            shape_name: Shape `cNvPr/@name` to remove (case-insensitive).

        Raises:
            FileNotFoundError: if the slide part does not exist.
        """

        wanted = shape_name.strip().casefold()
        slide_part = f"ppt/slides/slide{slide_number}.xml"

        if slide_part not in self.files:
            raise FileNotFoundError(f"Slide not found: {slide_part}")

        slide_root = etree.fromstring(self.files[slide_part])
        removed = False

        for shape in slide_root.xpath(".//p:sp", namespaces=NS):
            cNvPr = shape.xpath("./p:nvSpPr/p:cNvPr", namespaces=NS)
            if not cNvPr:
                continue

            actual = (cNvPr[0].get("name") or "").strip().casefold()
            if actual == wanted:
                shape.getparent().remove(shape)
                removed = True
                break

        if not removed:
            for frame in slide_root.xpath(".//p:graphicFrame", namespaces=NS):
                cNvPr = frame.xpath("./p:nvGraphicFramePr/p:cNvPr", namespaces=NS)
                if not cNvPr:
                    continue

                actual = (cNvPr[0].get("name") or "").strip().casefold()
                if actual == wanted:
                    frame.getparent().remove(frame)
                    removed = True
                    break

        if removed:
            self.files[slide_part] = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            self._has_changes = True

    def duplicate_slide(self, template_slide_number: int, before_section_name: str | None = None) -> int:
        source_slide_part = f"ppt/slides/slide{template_slide_number}.xml"

        if source_slide_part not in self.files:
            raise FileNotFoundError(f"Template slide not found: {source_slide_part}")

        """Duplicate a slide within the package.

        Args:
            template_slide_number: The slide number to copy.
            before_section_name: Optional section name to insert the new
                slide before; if omitted the slide is appended to the end.

        Returns:
            The inserted slide number (int).

        Raises:
            FileNotFoundError: if the template slide does not exist.
        """

        new_slide_number, new_slide_part = self._next_slide_part()

        self._deep_copy_part(source_slide_part, new_slide_part)
        self._add_content_type_override(new_slide_part, SLIDE_CONTENT_TYPE)

        inserted_slide_number = self._add_slide_to_presentation(slide_part=new_slide_part, before_section_name=before_section_name)

        self._has_changes = True
        return inserted_slide_number

    def duplicate_template_slide(self, template_name: str, before_section_name: str = "template_slides") -> "SlideProxy":
        """Duplicate a named template slide.

        This method looks for a shape named `TEMPLATE__<template_name>` on
        template slides (created in PowerPoint). The found slide is copied
        and the marker shape is removed from the new slide so it remains a
        proper generated slide.

        Args:
            template_name: Logical template identifier (without the `TEMPLATE__` prefix).
            before_section_name: Section name to insert generated slide before.

        Returns:
            A `SlideProxy` bound to the newly inserted slide.
        """

        marker_shape_name = f"TEMPLATE__{template_name}"

        template_slide_number = self.find_slide_by_shape_name(marker_shape_name)

        new_slide_number = self.duplicate_slide(template_slide_number=template_slide_number, before_section_name=before_section_name)

        # Remove marker from generated slide so future lookups only find template slides.
        self.remove_shape_on_slide(slide_number=new_slide_number, shape_name=marker_shape_name)

        # Return a bound SlideProxy for convenient chained edits.
        return self.get_slide(new_slide_number)

    def get_slide(self, slide_number: int) -> "SlideProxy":
        """Return a small helper proxy bound to `slide_number` so edit operations
        can be called directly on the returned object.

        Example:
            slide = editor.get_slide(5)
            slide.edit_textbox('Title', 'Hello')
        """
        return SlideProxy(self, slide_number)

    def drop_section(self, section_name: str, delete_slide_parts: bool = True) -> list[int]:
        """Remove a PowerPoint section and its slides from the presentation.

        Args:
            section_name: Section name to remove, case-insensitive.
            delete_slide_parts: When true, also delete the removed slide parts
                and any now-unreferenced private dependencies such as charts and
                embedded workbooks.

        Returns:
            The numeric slide part numbers removed from the presentation.
        """

        presentation_part = "ppt/presentation.xml"
        presentation_rels_part = "ppt/_rels/presentation.xml.rels"

        presentation_root = etree.fromstring(self.files[presentation_part])
        presentation_rels_root = etree.fromstring(self.files[presentation_rels_part])

        wanted = section_name.strip().casefold()
        sections = presentation_root.xpath(".//p14:section", namespaces=NS)
        target_section = None

        for section in sections:
            actual = (section.get("name") or "").strip().casefold()
            if actual == wanted:
                target_section = section
                break

        if target_section is None:
            raise ValueError(f"Section named '{section_name}' not found")

        section_slide_ids = {
            sld_id.get("id")
            for sld_id in target_section.xpath("./p14:sldIdLst/p14:sldId", namespaces=NS)
            if sld_id.get("id")
        }

        sld_id_lst = presentation_root.xpath("./p:sldIdLst", namespaces=NS)
        if not sld_id_lst:
            raise ValueError("No p:sldIdLst found in presentation.xml")

        rid_to_slide_part = {}
        for rel in presentation_rels_root.xpath("./pr:Relationship", namespaces=NS):
            if rel.get("Type") != REL_TYPE_SLIDE:
                continue

            rel_id = rel.get("Id")
            target = rel.get("Target")
            if rel_id and target:
                rid_to_slide_part[rel_id] = self._normalize_relationship_target(presentation_part, target)

        removed_slide_parts = []

        for sld_id in list(sld_id_lst[0].xpath("./p:sldId", namespaces=NS)):
            if sld_id.get("id") not in section_slide_ids:
                continue

            rel_id = sld_id.get(f"{{{R_NS}}}id")
            if rel_id and rel_id in rid_to_slide_part:
                removed_slide_parts.append(rid_to_slide_part[rel_id])

            sld_id.getparent().remove(sld_id)

            if rel_id:
                for rel in presentation_rels_root.xpath(f"./pr:Relationship[@Id='{rel_id}']", namespaces=NS):
                    rel.getparent().remove(rel)

        target_section.getparent().remove(target_section)

        self.files[presentation_part] = etree.tostring(presentation_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        self.files[presentation_rels_part] = etree.tostring(presentation_rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        if delete_slide_parts:
            for slide_part in removed_slide_parts:
                self._delete_part_and_unreferenced_targets(slide_part)

        self._has_changes = True

        removed_slide_numbers = []
        for slide_part in removed_slide_parts:
            match = SLIDE_RE.match(slide_part)
            if match:
                removed_slide_numbers.append(int(match.group(1)))

        return removed_slide_numbers
