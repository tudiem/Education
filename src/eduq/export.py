"""Xuất câu hỏi trong kho ra .docx và PDF.

.docx: ghép lại OOXML thô của từng câu (công thức MathType giữ nguyên dạng OLE + ảnh, hình giữ nguyên),
       dùng style/đánh số của file nguồn làm khuôn, bỏ đầu trang/watermark, đánh số lại "Câu n:" liền mạch.
PDF:   "chrome" = HTML + KaTeX in bằng Chrome headless (giống trang web);
       "libreoffice" = chuyển từ .docx (công thức hiện bằng ảnh preview gốc của MathType).
"""
from __future__ import annotations

import base64
import copy
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from sqlalchemy import select

from . import render
from .assets import AssetStore
from .db import Asset, Question, SourceFile
from .mathtype import TOOLS_DIR
from .package import NS, PARSER, DocxPackage, q
from .parse import QUESTION_LABEL_RE

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_PREFIX = "application/vnd.openxmlformats-officedocument.wordprocessingml."
CT_OVERRIDE = {
    "document": CT_PREFIX + "document.main+xml",
    "styles": CT_PREFIX + "styles+xml",
    "numbering": CT_PREFIX + "numbering+xml",
    "fontTable": CT_PREFIX + "fontTable+xml",
    "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
}
EXT_CT = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp",
    "wmf": "image/x-wmf", "emf": "image/x-emf", "tif": "image/tiff", "tiff": "image/tiff", "svg": "image/svg+xml",
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
W_NS = NS["w"]

DEFAULT_STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
<w:sz w:val="26"/><w:szCs w:val="26"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/>
<w:pPr><w:ind w:left="0"/></w:pPr></w:style></w:styles>""".encode()

DEFAULT_ABSTRACT = f"""<w:abstractNum xmlns:w="{W_NS}" w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="Câu %1:"/>
<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>
<w:rPr><w:b/><w:color w:val="0070C0"/></w:rPr></w:lvl></w:abstractNum>"""

DEFAULT_SECT = f"""<w:sectPr xmlns:w="{W_NS}"><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>"""


class ExportError(Exception):
    pass


@dataclass
class ExportResult:
    path: Path
    questions: int
    warnings: list[str] = field(default_factory=list)


# ------------------------------------------------------------------ chọn câu hỏi


def parse_codes(spec: str) -> list[str]:
    """"C0000001,C0000005-C0000008" → danh sách mã (giữ thứ tự, bỏ trùng)."""
    out: list[str] = []
    for part in (x.strip() for x in spec.split(",") if x.strip()):
        m = re.fullmatch(r"C(\d+)\s*-\s*C(\d+)", part, re.I)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if b < a or b - a > 5000:
                raise ExportError(f"Khoảng mã không hợp lệ: {part}")
            out += [f"C{n:07d}" for n in range(a, b + 1)]
        elif re.fullmatch(r"C\d{7}", part, re.I):
            out.append(part.upper())
        else:
            raise ExportError(f"Mã không hợp lệ: {part} (dạng C0000001)")
    return list(dict.fromkeys(out))


def select_questions(session, *, codes: list[str] | None = None, exam: str | None = None,
                     topic: str | None = None) -> list[Question]:
    from .db import Topic

    if codes:
        rows = {r.code: r for r in session.scalars(select(Question).where(Question.code.in_(codes)))}
        missing = [c for c in codes if c not in rows]
        if missing:
            raise ExportError(f"Không có trong kho: {', '.join(missing[:10])}")
        return [rows[c] for c in codes]
    stmt = select(Question).order_by(Question.id)
    if topic:
        stmt = stmt.join(Topic, Topic.id == Question.topic_id).where(Topic.name == topic)
    if exam:
        stmt = stmt.where(Question.exam == exam)
    rows = list(session.scalars(stmt))
    if not rows:
        raise ExportError("Không có câu hỏi nào khớp điều kiện chọn")
    return rows


# ------------------------------------------------------------------ khuôn (style, đánh số, trang)


@dataclass
class Template:
    styles: bytes
    theme: bytes | None
    font_table: bytes | None
    numbering: etree._Element
    q_abstract: etree._Element
    root_attrib: dict
    root_nsmap: dict
    sect_pr: etree._Element
    source_name: str | None = None


def _default_template() -> Template:
    return Template(
        styles=DEFAULT_STYLES, theme=None, font_table=None,
        numbering=etree.fromstring(f'<w:numbering xmlns:w="{W_NS}"/>'),
        q_abstract=etree.fromstring(DEFAULT_ABSTRACT), root_attrib={},
        root_nsmap={k: v for k, v in NS.items() if k in ("w", "r", "wp", "a", "pic", "v", "o", "m", "mc")},
        sect_pr=etree.fromstring(DEFAULT_SECT),
    )


def load_template(data: bytes | None, name: str | None = None) -> Template:
    if data is None:
        return _default_template()
    pkg = DocxPackage(data)
    by_kind = {r.type.rsplit("/", 1)[-1]: r for r in pkg.rels.values() if not r.external}

    def part(kind: str) -> bytes | None:
        r = by_kind.get(kind)
        return pkg.read(pkg.part_name(r)) if r and pkg.exists(pkg.part_name(r)) else None

    t = _default_template()
    root = pkg.document
    body = root.find(q("w", "body"))
    num_bytes = part("numbering")
    numbering = etree.fromstring(num_bytes, PARSER) if num_bytes else t.numbering

    # abstractNum của danh sách "Câu %1:" (mức 0 có số thứ tự) trong file nguồn
    q_abs = None
    for p in body.iter(q("w", "p")):
        npr = p.find(q("w", "pPr") + "/" + q("w", "numPr"))
        if npr is None or npr.find(q("w", "numId")) is None:
            continue
        nid = npr.find(q("w", "numId")).get(q("w", "val"))
        lv = npr.find(q("w", "ilvl"))
        if not pkg.numbering.is_ordered_top_level(nid, lv.get(q("w", "val")) if lv is not None else "0"):
            continue
        num = next((n for n in numbering.findall(q("w", "num")) if n.get(q("w", "numId")) == nid), None)
        ab_id = num.find(q("w", "abstractNumId")).get(q("w", "val")) if num is not None else None
        q_abs = next((a for a in numbering.findall(q("w", "abstractNum")) if a.get(q("w", "abstractNumId")) == ab_id), None)
        break

    sect = body.find(q("w", "sectPr"))
    sect = copy.deepcopy(sect) if sect is not None else t.sect_pr
    for tag in ("headerReference", "footerReference", "titlePg"):  # bỏ đầu/chân trang & watermark
        for e in sect.findall(q("w", tag)):
            sect.remove(e)
    return Template(
        styles=part("styles") or DEFAULT_STYLES, theme=part("theme"), font_table=part("fontTable"),
        numbering=numbering, q_abstract=copy.deepcopy(q_abs) if q_abs is not None else t.q_abstract,
        root_attrib=dict(root.attrib), root_nsmap=dict(root.nsmap), sect_pr=sect, source_name=name,
    )


def template_for(session, store: AssetStore, questions: list[Question]) -> tuple[Template, list[str]]:
    warnings: list[str] = []
    sf_ids = list(dict.fromkeys(qu.source_file_id for qu in questions))
    primary = session.get(SourceFile, sf_ids[0])
    if len(sf_ids) > 1:
        warnings.append(f"Câu hỏi đến từ {len(sf_ids)} file nguồn: dùng style của “{primary.filename}”, đánh số thống nhất")
    p = store.path(primary.sha256)
    if not p.exists():
        warnings.append("Không còn file nguồn trong kho (nhập từ phiên bản cũ): dùng style mặc định")
        return load_template(None), warnings
    return load_template(p.read_bytes(), primary.filename), warnings


A4_TWIPS = (11906, 16838)


def _apply_paper(sect, paper: str) -> tuple[int, int]:
    """paper="a4": đặt khổ A4 (giữ lề của nguồn); "source": giữ nguyên. Trả (rộng, cao) twips."""
    pg = sect.find(q("w", "pgSz"))
    if paper == "a4":
        if pg is None:
            pg = etree.Element(q("w", "pgSz"))
            typ = sect.find(q("w", "type"))
            if typ is not None:
                typ.addnext(pg)
            else:
                sect.insert(0, pg)
        pg.set(q("w", "w"), str(A4_TWIPS[0]))
        pg.set(q("w", "h"), str(A4_TWIPS[1]))
        pg.attrib.pop(q("w", "orient"), None)
    if pg is None:
        return A4_TWIPS
    return int(pg.get(q("w", "w"), A4_TWIPS[0])), int(pg.get(q("w", "h"), A4_TWIPS[1]))


def _add_question_numbering(tpl: Template) -> int:
    """Thêm một danh sách đánh số "Câu n:" vào numbering.xml; trả numId."""
    root = tpl.numbering
    abs_ids = [int(a.get(q("w", "abstractNumId"))) for a in root.findall(q("w", "abstractNum"))]
    num_ids = [int(n.get(q("w", "numId"))) for n in root.findall(q("w", "num"))]
    new_abs, new_num = max(abs_ids + [-1]) + 1, max(num_ids + [0]) + 1
    ab = copy.deepcopy(tpl.q_abstract)
    ab.set(q("w", "abstractNumId"), str(new_abs))
    for tag in ("nsid", "tmpl"):
        for e in ab.findall(q("w", tag)):
            ab.remove(e)
    for ps in ab.iter(q("w", "pStyle")):  # không để danh sách gắn vào style có sẵn
        ps.getparent().remove(ps)
    existing = root.findall(q("w", "abstractNum"))
    if existing:
        existing[-1].addnext(ab)
    else:
        root.insert(0, ab)
    num = etree.Element(q("w", "num"))
    num.set(q("w", "numId"), str(new_num))
    etree.SubElement(num, q("w", "abstractNumId")).set(q("w", "val"), str(new_abs))
    nums = root.findall(q("w", "num"))
    (nums[-1] if nums else ab).addnext(num)
    return new_num


# ------------------------------------------------------------------ dựng gói .docx


class _Out:
    """Gói .docx đầu ra: các part, quan hệ, kiểu nội dung."""

    def __init__(self):
        self.parts: dict[str, bytes] = {}
        self.rels: list[tuple[str, str, str, bool]] = []  # (id, type, target, external)
        self._n = 0
        self._asset_rid: dict[tuple[str, str], str] = {}
        self._counters = {"image": 0, "oleObject": 0}

    def rid(self) -> str:
        self._n += 1
        return f"rId{self._n}"

    def add_part_rel(self, kind: str, target: str, data: bytes) -> None:
        self.parts["word/" + target] = data
        self.rels.append((self.rid(), REL + kind, target, False))

    def add_external(self, kind: str, target: str) -> str:
        for r in self.rels:
            if r[3] and r[2] == target and r[1] == REL + kind:
                return r[0]
        rid = self.rid()
        self.rels.append((rid, REL + kind, target, True))
        return rid

    def add_asset(self, kind: str, sha: str, ext: str, data: bytes) -> str:
        key = (kind, sha)
        if key in self._asset_rid:
            return self._asset_rid[key]
        self._counters[kind] = self._counters.get(kind, 0) + 1
        if kind == "oleObject":
            target = f"embeddings/oleObject{self._counters[kind]}.bin"
        else:
            target = f"media/image{self._counters[kind]}.{ext}"
        self.parts["word/" + target] = data
        rid = self.rid()
        self.rels.append((rid, REL + kind, target, False))
        self._asset_rid[key] = rid
        return rid


def _el_text(parts_el) -> str:
    return "".join(t.text or "" for t in parts_el.iter(q("w", "t")))


def _strip_label(p) -> bool:
    """Bỏ nhãn gõ tay "Câu N." ở đầu đoạn (kể cả khi bị tách thành nhiều run). True nếu đã bỏ."""
    ts = list(p.iter(q("w", "t")))
    m = QUESTION_LABEL_RE.match("".join(t.text or "" for t in ts))
    if not m:
        return False
    left = m.end()
    for t in ts:
        if left <= 0:
            break
        txt = t.text or ""
        cut = min(left, len(txt))
        t.text = txt[cut:]
        t.set(XML_SPACE, "preserve")
        left -= cut
    return True


def _set_numbering(p, num_id: int) -> None:
    ppr = p.find(q("w", "pPr"))
    if ppr is None:
        ppr = etree.Element(q("w", "pPr"))
        p.insert(0, ppr)
    for old in ppr.findall(q("w", "numPr")):
        ppr.remove(old)
    numpr = etree.Element(q("w", "numPr"))
    etree.SubElement(numpr, q("w", "ilvl")).set(q("w", "val"), "0")
    etree.SubElement(numpr, q("w", "numId")).set(q("w", "val"), str(num_id))
    idx = 0
    for i, ch in enumerate(ppr):  # numPr phải đứng sau các phần tử này theo schema
        if isinstance(ch.tag, str) and etree.QName(ch).localname in (
            "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"
        ):
            idx = i + 1
    ppr.insert(idx, numpr)


def _remap_rels(els, raw_rels: dict, out: _Out, store: AssetStore, code: str, warnings: list[str]) -> None:
    mapping: dict[str, str] = {}
    for rid, e in raw_rels.items():
        kind = e.get("type", "")
        if e.get("external"):
            mapping[rid] = out.add_external(kind, e["target"])
        elif e.get("asset"):
            sha = e["asset"]
            if not store.path(sha).exists():
                warnings.append(f"{code}: thiếu tài sản {sha[:8]} trong kho")
                continue
            ext = e["target"].rsplit(".", 1)[-1].lower() if "." in e["target"] else "bin"
            mapping[rid] = out.add_asset("oleObject" if kind == "oleObject" else "image", sha, ext, store.read(sha))
        else:
            warnings.append(f"{code}: quan hệ {rid} ({kind}) chưa hỗ trợ khi xuất")
    rattr = "{%s}" % NS["r"]
    for el in els:
        for node in el.iter():
            if not isinstance(node.tag, str):
                continue
            for k in [k for k in node.attrib if k.startswith(rattr)]:
                v = node.attrib[k]
                if v in mapping:
                    node.attrib[k] = mapping[v]
                else:
                    del node.attrib[k]  # tránh r:id treo làm Word báo lỗi file


def export_docx(questions: list[Question], tpl: Template, store: AssetStore, out_path: str | Path,
                title: str | None = None, paper: str = "a4") -> ExportResult:
    if not questions:
        raise ExportError("Không có câu hỏi để xuất")
    out = _Out()
    warnings: list[str] = []
    num_id = _add_question_numbering(tpl)

    doc = etree.Element(q("w", "document"), attrib=dict(tpl.root_attrib), nsmap=dict(tpl.root_nsmap))
    body = etree.SubElement(doc, q("w", "body"))
    if title:
        tp = etree.SubElement(body, q("w", "p"))
        ppr = etree.SubElement(tp, q("w", "pPr"))
        etree.SubElement(ppr, q("w", "spacing")).set(q("w", "after"), "240")
        etree.SubElement(ppr, q("w", "jc")).set(q("w", "val"), "center")
        rpr = etree.SubElement(etree.SubElement(tp, q("w", "r")), q("w", "rPr"))
        etree.SubElement(rpr, q("w", "b"))
        etree.SubElement(rpr, q("w", "sz")).set(q("w", "val"), "30")
        etree.SubElement(tp[-1], q("w", "t")).text = title

    for qu in questions:
        els = list(etree.fromstring(f"<root>{qu.raw_ooxml}</root>", PARSER))
        for el in els:
            for sp in list(el.iter(q("w", "sectPr"))):  # ngắt section của file nguồn
                sp.getparent().remove(sp)
        first = els[0]
        if first.tag == q("w", "p"):
            _strip_label(first)
            _set_numbering(first, num_id)
        else:
            warnings.append(f"{qu.code}: câu bắt đầu bằng bảng, không đánh số được")
        _remap_rels(els, qu.raw_rels, out, store, qu.code, warnings)
        for el in els:
            body.append(el)

    for i, d in enumerate(body.iter("{%s}docPr" % NS["wp"]), 1):  # id hình phải duy nhất toàn tài liệu
        d.set("id", str(i))
    sect = copy.deepcopy(tpl.sect_pr)
    _apply_paper(sect, paper)
    body.append(sect)

    # --- các part cố định
    out.parts["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    out.add_part_rel("styles", "styles.xml", tpl.styles)
    out.add_part_rel("numbering", "numbering.xml",
                     etree.tostring(tpl.numbering, xml_declaration=True, encoding="UTF-8", standalone=True))
    overrides = {"/word/document.xml": CT_OVERRIDE["document"], "/word/styles.xml": CT_OVERRIDE["styles"],
                 "/word/numbering.xml": CT_OVERRIDE["numbering"]}
    if tpl.font_table:
        out.add_part_rel("fontTable", "fontTable.xml", tpl.font_table)
        overrides["/word/fontTable.xml"] = CT_OVERRIDE["fontTable"]
    if tpl.theme:
        out.add_part_rel("theme", "theme/theme1.xml", tpl.theme)
        overrides["/word/theme/theme1.xml"] = CT_OVERRIDE["theme"]

    rels = etree.Element("Relationships", nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"})
    for rid, typ, target, ext in out.rels:
        r = etree.SubElement(rels, "Relationship", Id=rid, Type=typ, Target=target)
        if ext:
            r.set("TargetMode", "External")
    out.parts["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    out.parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
        f'Type="{REL}officeDocument" Target="word/document.xml"/></Relationships>'
    ).encode()

    exts = sorted({n.rsplit(".", 1)[-1].lower() for n in out.parts if n.startswith(("word/media/", "word/embeddings/"))})
    ct = etree.Element("Types", nsmap={None: "http://schemas.openxmlformats.org/package/2006/content-types"})
    etree.SubElement(ct, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    etree.SubElement(ct, "Default", Extension="xml", ContentType="application/xml")
    for e in exts:
        etree.SubElement(ct, "Default", Extension=e, ContentType=EXT_CT.get(e, "application/octet-stream"))
    for part, typ in overrides.items():
        etree.SubElement(ct, "Override", PartName=part, ContentType=typ)
    out.parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    order = ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels"]
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order + sorted(n for n in out.parts if n not in order):
            z.writestr(name, out.parts[name])

    warnings += [f"kiểm tra .docx: {p}" for p in check_docx(out_path)]
    return ExportResult(out_path, len(questions), warnings)


# ------------------------------------------------------------------ kiểm tra cấu trúc .docx


def check_docx(path: str | Path) -> list[str]:
    """Kiểm tra các lỗi hay làm Word báo "file bị hỏng": XML sai, r:id treo, thiếu part/kiểu nội dung, id trùng."""
    problems: list[str] = []
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return ["không phải file zip"]
    names = set(z.namelist())
    docs: dict[str, etree._Element] = {}
    for n in names:
        if n.endswith((".xml", ".rels")):
            try:
                docs[n] = etree.fromstring(z.read(n), PARSER)
            except etree.XMLSyntaxError as e:
                problems.append(f"XML hỏng: {n}: {e}")
    if "word/document.xml" not in docs or "word/_rels/document.xml.rels" not in docs:
        return problems + ["thiếu document.xml hoặc rels"]

    rel_ids = {}
    for r in docs["word/_rels/document.xml.rels"]:
        rel_ids[r.get("Id")] = r
        if r.get("TargetMode") != "External" and "word/" + r.get("Target") not in names:
            problems.append(f"rels trỏ tới part không tồn tại: {r.get('Target')}")
    rattr = "{%s}" % NS["r"]
    for node in docs["word/document.xml"].iter():
        if isinstance(node.tag, str):
            for k, v in node.attrib.items():
                if k.startswith(rattr) and v not in rel_ids:
                    problems.append(f"r:id treo: {v}")
    ids = [d.get("id") for d in docs["word/document.xml"].iter("{%s}docPr" % NS["wp"])]
    if len(ids) != len(set(ids)):
        problems.append("wp:docPr id trùng")

    ct = docs.get("[Content_Types].xml")
    defaults = {e.get("Extension").lower() for e in ct if etree.QName(e).localname == "Default"} if ct is not None else set()
    overrides = {e.get("PartName") for e in ct if etree.QName(e).localname == "Override"} if ct is not None else set()
    for n in names:
        if n == "[Content_Types].xml":
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        if ext not in defaults and "/" + n not in overrides:
            problems.append(f"thiếu content type cho {n}")

    nums = docs.get("word/numbering.xml")
    have = {n.get(q("w", "numId")) for n in nums.findall(q("w", "num"))} if nums is not None else set()
    for nid in docs["word/document.xml"].iter(q("w", "numId")):
        if nid.get(q("w", "val")) not in have:
            problems.append(f"numId {nid.get(q('w', 'val'))} không có trong numbering.xml")
            break
    return problems


# ------------------------------------------------------------------ PDF


def _chrome_binary() -> str | None:
    env = os.environ.get("EDUQ_CHROME")
    if env and shutil.which(env):
        return shutil.which(env)
    for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome"):
        if shutil.which(name):
            return shutil.which(name)
    return None


def _soffice_binary() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def pdf_engines() -> list[str]:
    out = []
    if _chrome_binary() and (TOOLS_DIR / "node_modules" / "katex" / "dist").is_dir():
        out.append("chrome")
    if _soffice_binary():
        out.append("libreoffice")
    return out


PRINT_CSS = (
    "body{font-size:13pt;line-height:1.5;max-width:none;margin:0;padding:0}"
    "h1{font-size:16pt;text-align:center;margin:0 0 12px}.q{margin:8px 0;break-inside:avoid}"
    ".q:after{content:'';display:block;clear:both}"
)


def _exam_html(questions: list[Question], asset_url, title: str | None, page_in: tuple[float, float]) -> str:
    parts = [f"<h1>{render.escape(title)}</h1>"] if title else []
    for i, qu in enumerate(questions, 1):
        parts.append(f'<div class="q">{render.question_html(qu.content, asset_url, label=f"Câu {i}:")}</div>')
    css = f"@page{{size:{page_in[0]:.3f}in {page_in[1]:.3f}in;margin:16mm 15mm}}" + PRINT_CSS
    return render.page_html("\n".join(parts), title=title or "Đề", extra_css=css)


def _pdf_chrome(questions, session, store, out_path: Path, title: str | None, page_in: tuple[float, float]) -> None:
    chrome = _chrome_binary()
    kdist = TOOLS_DIR / "node_modules" / "katex" / "dist"
    if not chrome or not kdist.is_dir():
        raise ExportError("Thiếu Chrome/Chromium hoặc KaTeX (chạy `npm install` trong tools/)")
    mimes: dict[str, str] = {}

    def asset_url(sha: str) -> str:
        if sha not in mimes:
            a = session.get(Asset, sha)
            mimes[sha] = a.mime if a else "application/octet-stream"
        return f"data:{mimes[sha]};base64,{base64.b64encode(store.read(sha)).decode()}"

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        shutil.copytree(kdist, td / "katex")
        (td / "index.html").write_text(_exam_html(questions, asset_url, title, page_in), encoding="utf-8")
        cmd = [chrome, "--headless", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=15000",
               f"--user-data-dir={td / 'profile'}", f"--print-to-pdf={td / 'out.pdf'}"]
        if hasattr(os, "geteuid") and os.geteuid() == 0:
            cmd.append("--no-sandbox")  # Chrome không chạy được dưới root nếu còn sandbox (môi trường container)
        r = subprocess.run(cmd + [(td / "index.html").as_uri()], capture_output=True, text=True, timeout=180)
        if not (td / "out.pdf").exists():
            raise ExportError(f"Chrome không tạo được PDF: {r.stderr[-300:]}")
        shutil.copyfile(td / "out.pdf", out_path)


def _pdf_libreoffice(docx_path: Path, out_path: Path) -> None:
    soffice = _soffice_binary()
    if not soffice:
        raise ExportError("Không tìm thấy LibreOffice (soffice)")
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run(
            [soffice, f"-env:UserInstallation={Path(td).as_uri()}/lo", "--headless", "--convert-to", "pdf",
             "--outdir", td, str(docx_path)], capture_output=True, text=True, timeout=300,
        )
        pdf = Path(td) / (docx_path.stem + ".pdf")
        if not pdf.exists():
            raise ExportError(f"LibreOffice không tạo được PDF: {r.stderr[-300:]}")
        shutil.copyfile(pdf, out_path)


def export_pdf(questions: list[Question], session, store: AssetStore, out_path: str | Path,
               title: str | None = None, engine: str = "auto", paper: str = "a4") -> ExportResult:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    avail = pdf_engines()
    if engine == "auto":
        if not avail:
            raise ExportError("Không có công cụ tạo PDF: cài Chrome/Chromium (+ `npm install` trong tools/) hoặc LibreOffice")
        engine = avail[0]
    warnings: list[str] = []
    if engine == "chrome":
        tpl, w = template_for(session, store, questions)
        warnings += w
        wt, ht = _apply_paper(copy.deepcopy(tpl.sect_pr), paper)
        _pdf_chrome(questions, session, store, out_path, title, (wt / 1440, ht / 1440))
        for qu in questions:
            if qu.status != "ok":
                warnings.append(f"{qu.code}: có công thức cần review, PDF có thể thiếu/sai phần đó")
    elif engine == "libreoffice":
        with tempfile.TemporaryDirectory() as td:
            tpl, w = template_for(session, store, questions)
            res = export_docx(questions, tpl, store, Path(td) / out_path.with_suffix(".docx").name, title, paper)
            warnings += w + res.warnings
            _pdf_libreoffice(res.path, out_path)
    else:
        raise ExportError(f"Engine PDF không hợp lệ: {engine}")
    return ExportResult(out_path, len(questions), warnings)
