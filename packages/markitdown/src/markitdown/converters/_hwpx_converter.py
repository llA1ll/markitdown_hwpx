import io
import re
import zipfile

from pathlib import Path
from typing import BinaryIO, Any

from defusedxml import ElementTree as DET

from .._base_converter import DocumentConverter, DocumentConverterResult
from .._stream_info import StreamInfo

ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/hwp+zip",
    "application/x-hwpx",
    "application/octet-stream",
]

ACCEPTED_FILE_EXTENSIONS = [".hwpx"]


class HwpxConverter(DocumentConverter):
    """
    Converts HWPX files to Markdown.

    HWPX is a ZIP container of XML sections. This converter extracts section text,
    simple tables, equations, and image references into a linear Markdown form.
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> DocumentConverterResult:
        hwpx_bytes = io.BytesIO(file_stream.read())
        with zipfile.ZipFile(hwpx_bytes) as archive:
            section_names = self._list_section_names(archive)
            markdown_chunks: list[str] = []

            for section_name in section_names:
                root = DET.fromstring(archive.read(section_name))
                section_markdown = self._convert_section(root)
                if section_markdown.strip():
                    markdown_chunks.append(section_markdown.strip())

        title = None
        if markdown_chunks:
            first_nonempty = next(
                (line.strip() for line in markdown_chunks[0].splitlines() if line.strip()),
                None,
            )
            title = first_nonempty

        return DocumentConverterResult(
            markdown="\n\n".join(markdown_chunks).strip(),
            title=title,
        )

    def _list_section_names(self, archive: zipfile.ZipFile) -> list[str]:
        section_names = [
            name
            for name in archive.namelist()
            if name.lower().startswith("contents/section")
            and name.lower().endswith(".xml")
        ]

        def sort_key(name: str) -> tuple[int, str]:
            match = re.search(r"section(\d+)\.xml$", name.lower())
            if match:
                return (int(match.group(1)), name.lower())
            return (10**9, name.lower())

        return sorted(section_names, key=sort_key)

    def _convert_section(self, root: Any) -> str:
        blocks: list[str] = []

        for paragraph in root.iter():
            if not self._tag_name(paragraph).endswith("p"):
                continue

            paragraph_lines = self._paragraph_to_markdown(paragraph)
            if paragraph_lines:
                blocks.extend(paragraph_lines)
                blocks.append("")

        while blocks and not blocks[-1].strip():
            blocks.pop()
        return "\n".join(blocks)

    def _paragraph_to_markdown(self, paragraph: Any) -> list[str]:
        inline_parts: list[str] = []
        block_lines: list[str] = []

        for node in paragraph.iter():
            tag = self._tag_name(node)
            if tag == "t":
                text = self._normalize_text(node.text or "")
                if text:
                    inline_parts.append(text)
            elif tag == "script":
                script = self._normalize_text(node.text or "")
                if script:
                    inline_parts.append(f"${script}$")
            elif tag == "img":
                ref = self._image_ref(node)
                if ref:
                    continue
            elif tag == "tbl":
                if inline_parts:
                    block_lines.append(self._join_inline(inline_parts))
                    inline_parts = []
                table_md = self._table_to_markdown(node)
                if table_md:
                    block_lines.append(table_md)

        if inline_parts:
            block_lines.append(self._join_inline(inline_parts))

        return [line for line in block_lines if line.strip()]

    def _table_to_markdown(self, tbl_node: Any) -> str:
        rows: list[list[str]] = []
        for tr in list(tbl_node):
            if self._tag_name(tr) != "tr":
                continue
            row: list[str] = []
            for tc in list(tr):
                if self._tag_name(tc) != "tc":
                    continue
                cell_parts: list[str] = []
                for sub in tc.iter():
                    tag = self._tag_name(sub)
                    if tag == "t":
                        text = self._normalize_text(sub.text or "")
                        if text:
                            cell_parts.append(text)
                    elif tag == "script":
                        script = self._normalize_text(sub.text or "")
                        if script:
                            cell_parts.append(f"${script}$")
                row.append(self._join_inline(cell_parts))
            if any(cell.strip() for cell in row):
                rows.append(row)

        if not rows:
            return ""

        column_count = max(len(row) for row in rows)
        normalized_rows = [row + [""] * (column_count - len(row)) for row in rows]

        header = normalized_rows[0]
        separator = ["---"] * column_count
        body = normalized_rows[1:]

        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(separator) + " |",
        ]
        lines.extend("| " + " | ".join(row) + " |" for row in body)
        return "\n".join(lines)

    def _image_ref(self, node: Any) -> str:
        return (
            node.attrib.get("binaryItemIDRef")
            or node.attrib.get(
                "{http://www.hancom.co.kr/hwpml/2011/ctrl}binaryItemIDRef", ""
            )
            or ""
        ).strip()

    def _join_inline(self, parts: list[str]) -> str:
        return self._normalize_text(" ".join(parts))

    def _normalize_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()

    def _tag_name(self, node: Any) -> str:
        tag = getattr(node, "tag", "")
        if "}" in tag:
            return tag.split("}", 1)[1]
        return tag
