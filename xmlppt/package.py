from zipfile import ZipFile
from lxml import etree
import msvcrt
import os
import posixpath
import re
import win32con
import win32file
from .constants import *


class PackageMixin:
    @staticmethod
    def _normalize_relationship_target(base_part: str, target: str) -> str:
        base_dir = posixpath.dirname(base_part)
        return posixpath.normpath(posixpath.join(base_dir, target))

    def _load_pptx_files(self, pptx_path: str) -> dict[str, bytes]:
        if os.name == "nt":
            try:
                return self._load_pptx_files_windows_shared(pptx_path)
            except Exception:
                pass

        with ZipFile(pptx_path, "r") as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    def _load_pptx_files_windows_shared(self, pptx_path: str) -> dict[str, bytes]:
        """Open a PPTX with Windows shared read access so the file can be read
        even when PowerPoint has it open.
        """
        handle = win32file.CreateFile(
            os.path.abspath(pptx_path),
            win32file.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )

        try:
            fd = msvcrt.open_osfhandle(handle.Detach(), os.O_RDONLY)
        except Exception:
            handle.Close()
            raise

        with os.fdopen(fd, "rb") as shared_file:
            with ZipFile(shared_file, "r") as archive:
                return {name: archive.read(name) for name in archive.namelist()}

    def _write_pptx_files(self, output_pptx: str) -> None:
        with ZipFile(output_pptx, "w") as archive:
            for name, data in self.files.items():
                archive.writestr(name, data)

    def _default_output_name(self, input_pptx: str) -> str:
        base, ext = os.path.splitext(input_pptx)
        return f"{base}_updated{ext}"

    def save(self, output_pptx: str | None = None) -> str:
        """Write the in-memory PPTX package back to disk.

        Args:
            output_pptx: Optional output filename. If omitted, a new file
                is created by appending `_updated` before the extension of
                the input file.

        Returns:
            The path to the written output file.
        """

        if output_pptx is None:
            output_pptx = self._default_output_name(self.input_pptx)

        self._write_pptx_files(output_pptx)
        self._has_changes = False
        return output_pptx

    # Package / part helpers
    def _rels_part_for(self, part_name: str) -> str:
        folder = posixpath.dirname(part_name)
        filename = posixpath.basename(part_name)
        return f"{folder}/_rels/{filename}.rels"

    def _next_numbered_part_name(self, folder: str, prefix: str, ext: str) -> str:
        pattern = re.compile(rf"^{re.escape(folder)}/{re.escape(prefix)}(\d+){re.escape(ext)}$")

        used = []
        for name in self.files:
            match = pattern.match(name)
            if match:
                used.append(int(match.group(1)))

        next_num = max(used, default=0) + 1
        return f"{folder}/{prefix}{next_num}{ext}"

    def _next_slide_part(self) -> tuple[int, str]:
        slide_numbers = [
            int(match.group(1))
            for name in self.files
            if (match := SLIDE_RE.match(name))
        ]

        next_num = max(slide_numbers, default=0) + 1
        return next_num, f"ppt/slides/slide{next_num}.xml"

    def _next_relationship_id(self, rels_root) -> str:
        used = []

        for rel in rels_root.xpath("./pr:Relationship", namespaces=NS):
            rid = rel.get("Id", "")
            match = re.match(r"rId(\d+)$", rid)
            if match:
                used.append(int(match.group(1)))

        return f"rId{max(used, default=0) + 1}"

    def _next_slide_id(self, presentation_root) -> str:
        ids = []

        for sld_id in presentation_root.xpath(".//p:sldId", namespaces=NS):
            value = sld_id.get("id")
            if value and value.isdigit():
                ids.append(int(value))

        return str(max(ids, default=255) + 1)

    def _add_content_type_override(self, part_name: str, content_type: str) -> None:
        content_types_part = "[Content_Types].xml"
        root = etree.fromstring(self.files[content_types_part])

        part_name_with_slash = f"/{part_name}"

        existing = root.xpath(
            f"./ct:Override[@PartName='{part_name_with_slash}']",
            namespaces=NS,
        )

        if existing:
            existing[0].set("ContentType", content_type)
        else:
            override = etree.SubElement(root, f"{{{CT_NS}}}Override")
            override.set("PartName", part_name_with_slash)
            override.set("ContentType", content_type)

        self.files[content_types_part] = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

    def _remove_content_type_override(self, part_name: str) -> None:
        content_types_part = "[Content_Types].xml"
        if content_types_part not in self.files:
            return

        root = etree.fromstring(self.files[content_types_part])
        part_name_with_slash = f"/{part_name}"

        for override in root.xpath(f"./ct:Override[@PartName='{part_name_with_slash}']", namespaces=NS):
            override.getparent().remove(override)

        self.files[content_types_part] = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

    def _internal_relationship_targets(self, part_name: str) -> list[str]:
        rels_part = self._rels_part_for(part_name)
        if rels_part not in self.files:
            return []

        rels_root = etree.fromstring(self.files[rels_part])
        targets = []

        for rel in rels_root.xpath("./pr:Relationship", namespaces=NS):
            if rel.get("TargetMode") == "External":
                continue

            target = rel.get("Target")
            if not target:
                continue

            targets.append(self._normalize_relationship_target(part_name, target))

        return targets

    def _part_is_referenced(self, part_name: str) -> bool:
        for rels_part, data in self.files.items():
            if not rels_part.endswith(".rels"):
                continue

            if rels_part == "_rels/.rels":
                source_part = ""
            else:
                rels_folder, rels_filename = rels_part.rsplit("/_rels/", 1)
                source_part = f"{rels_folder}/{rels_filename[:-5]}"

            rels_root = etree.fromstring(data)

            for rel in rels_root.xpath("./pr:Relationship", namespaces=NS):
                if rel.get("TargetMode") == "External":
                    continue

                target = rel.get("Target")
                if not target:
                    continue

                if self._normalize_relationship_target(source_part, target) == part_name:
                    return True

        return False

    def _delete_part_and_unreferenced_targets(self, part_name: str) -> None:
        related_parts = self._internal_relationship_targets(part_name)

        self.files.pop(part_name, None)
        self.files.pop(self._rels_part_for(part_name), None)
        self._remove_content_type_override(part_name)

        for related_part in related_parts:
            if related_part in self.files and not self._part_is_referenced(related_part):
                self._delete_part_and_unreferenced_targets(related_part)
