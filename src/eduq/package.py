"""Đọc gói .docx (zip OOXML): các part, quan hệ (rels), đánh số."""
from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass
from functools import cached_property
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def q(prefix: str, name: str) -> str:
    """Tên có namespace dạng Clark: q('w','p') -> '{...}p'."""
    return "{%s}%s" % (NS[prefix], name)


# Chống XXE / tải tài nguyên ngoài khi đọc file người dùng upload.
PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
_PARSER = PARSER

MAX_UNCOMPRESSED_BYTES = 500 * 1024 * 1024


class DocxError(Exception):
    """File không phải .docx hợp lệ hoặc không đọc được."""


@dataclass(frozen=True)
class Rel:
    id: str
    type: str
    target: str
    external: bool


class DocxPackage:
    def __init__(self, source: str | Path | bytes):
        fp = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else str(source)
        try:
            self._zip = zipfile.ZipFile(fp)
        except (zipfile.BadZipFile, FileNotFoundError, IsADirectoryError) as e:
            raise DocxError(f"Không mở được file .docx: {e}") from e
        total = sum(i.file_size for i in self._zip.infolist())
        if total > MAX_UNCOMPRESSED_BYTES:
            raise DocxError("File giải nén quá lớn")
        if "word/document.xml" not in self._zip.namelist():
            raise DocxError("Thiếu word/document.xml: không phải file Word (.docx)")

    def read(self, name: str) -> bytes:
        try:
            return self._zip.read(name)
        except KeyError as e:
            raise DocxError(f"Thiếu part {name}") from e

    def exists(self, name: str) -> bool:
        return name in self._zip.namelist()

    def xml(self, name: str) -> etree._Element:
        return etree.fromstring(self.read(name), _PARSER)

    @cached_property
    def document(self) -> etree._Element:
        return self.xml("word/document.xml")

    @cached_property
    def rels(self) -> dict[str, Rel]:
        name = "word/_rels/document.xml.rels"
        if not self.exists(name):
            return {}
        out = {}
        for r in self.xml(name):
            out[r.get("Id")] = Rel(
                id=r.get("Id"),
                type=r.get("Type", ""),
                target=r.get("Target", ""),
                external=r.get("TargetMode") == "External",
            )
        return out

    def part_name(self, rel: Rel) -> str:
        """Tên part trong zip của một quan hệ nội bộ (tương đối so với word/)."""
        t = rel.target.lstrip("/") if rel.target.startswith("/") else posixpath.join("word", rel.target)
        return posixpath.normpath(t)

    @cached_property
    def numbering(self) -> "Numbering":
        if self.exists("word/numbering.xml"):
            return Numbering(self.xml("word/numbering.xml"))
        return Numbering(None)


class Numbering:
    """Tra cứu danh sách đánh số: một mức có phải là số thứ tự (không phải bullet)."""

    def __init__(self, root: etree._Element | None):
        self._num_to_abs: dict[str, str] = {}
        self._lvl: dict[tuple[str, str], tuple[str, str]] = {}
        if root is None:
            return
        for a in root.findall(q("w", "abstractNum")):
            aid = a.get(q("w", "abstractNumId"))
            for lvl in a.findall(q("w", "lvl")):
                ilvl = lvl.get(q("w", "ilvl"))
                fmt = lvl.find(q("w", "numFmt"))
                txt = lvl.find(q("w", "lvlText"))
                self._lvl[(aid, ilvl)] = (
                    fmt.get(q("w", "val")) if fmt is not None else "",
                    txt.get(q("w", "val")) if txt is not None else "",
                )
        for n in root.findall(q("w", "num")):
            ab = n.find(q("w", "abstractNumId"))
            if ab is not None:
                self._num_to_abs[n.get(q("w", "numId"))] = ab.get(q("w", "val"))

    def is_ordered_top_level(self, num_id: str, ilvl: str = "0") -> bool:
        if ilvl != "0":
            return False
        aid = self._num_to_abs.get(num_id)
        fmt, text = self._lvl.get((aid, ilvl), ("", ""))
        return fmt not in ("bullet", "none", "") and "%" in text
