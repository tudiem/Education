"""Cắt tài liệu Word thành đề → câu hỏi và trích nội dung (chữ, công thức, hình, bảng).

Nhận diện câu hỏi: đoạn là một mục của danh sách đánh số (vd. "Câu %1:") HOẶC bắt đầu bằng
"Câu N." / "Bài N". Các đoạn không đánh số theo sau thuộc về câu đó (vd. dòng A. B. C. D.).
Tiêu đề đề thi ("ĐỀ SỐ n") và phần ("I. TRẮC NGHIỆM…") được ghi lại thành metadata.
Mỗi câu giữ nguyên OOXML thô để xuất lại không mất gì.
"""
from __future__ import annotations

import hashlib
import posixpath
import re
from dataclasses import dataclass, field

from lxml import etree

from .package import NS, DocxPackage, q

EXAM_RE = re.compile(r"^\s*đề\s+(?:thi\s+)?(?:số\s+)?\d+\b", re.I)
SECTION_RE = re.compile(r"^\s*(?:I{1,3}|IV|V)\.\s+\S")
QUESTION_RE = re.compile(r"^\s*(?:câu|bài)\s*\d+\s*[.:)]", re.I)
QUESTION_LABEL_RE = re.compile(r"^\s*(?:câu|bài)\s*\d+\s*[.:)]\s*", re.I)
POINTS_RE = re.compile(r"\((\d+(?:[.,]\d+)?)\s*điểm\)", re.I)
OPTION_RE = re.compile(r"(?:(?<=\s)|^)([A-D])\.(?=\s|$)")

GEOM_RE = re.compile(
    r"đường thẳng|góc|tam giác|hình vẽ|\btia\b|đỉnh|song song|vuông góc|đoạn thẳng|hình (?:thang|bình hành|chữ nhật|vuông|thoi)",
    re.I,
)
MOVED_FIGURE = "figure_moved_from_previous_question"

MIME = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
    "bmp": "image/bmp", "wmf": "image/x-wmf", "emf": "image/x-emf", "svg": "image/svg+xml",
    "tif": "image/tiff", "tiff": "image/tiff", "bin": "application/octet-stream",
}


@dataclass
class AssetBlob:
    sha256: str
    kind: str  # "image" | "ole" | "preview"
    mime: str
    name: str
    data: bytes


@dataclass
class ParsedQuestion:
    exam: str | None
    section: str | None
    number: int  # thứ tự trong đề
    qtype: str  # "multiple_choice" | "free_response"
    points: float | None
    content: dict  # {"stem": [block...], "options": [{"label","inlines"}...] | None}
    raw_ooxml: str
    rel_ids: set[str]
    issues: list[str] = field(default_factory=list)


@dataclass
class ParsedDocument:
    questions: list[ParsedQuestion]
    assets: dict[str, AssetBlob]  # sha256 → blob (ảnh, OLE, preview)
    ole_blobs: dict[str, bytes]  # sha256 → bytes OLE cần chuyển đổi công thức
    issues: list[str] = field(default_factory=list)


# ------------------------------------------------------------------ tiện ích


def _local(el) -> str:
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def _ns(el) -> str:
    return etree.QName(el).namespace if isinstance(el.tag, str) else ""


def para_text(p) -> str:
    """Chỉ lấy chữ trong w:t (không lấy số offset của hình nổi như wp:posOffset)."""
    return "".join(t.text or "" for t in p.iter(q("w", "t")))


def _on(el, name: str) -> bool:
    """Thuộc tính bật/tắt kiểu w:b: có phần tử và không phải val=0/false."""
    e = el.find(q("w", name))
    if e is None:
        return False
    return e.get(q("w", "val"), "1").lower() not in ("0", "false", "off")


def _marks(rpr) -> dict:
    if rpr is None:
        return {}
    m: dict = {}
    if _on(rpr, "b"):
        m["b"] = True
    if _on(rpr, "i"):
        m["i"] = True
    u = rpr.find(q("w", "u"))
    if u is not None and u.get(q("w", "val"), "single") not in ("none",):
        m["u"] = True
    va = rpr.find(q("w", "vertAlign"))
    if va is not None and va.get(q("w", "val")) in ("superscript", "subscript"):
        m["sup" if va.get(q("w", "val")) == "superscript" else "sub"] = True
    return m


class _Ctx:
    def __init__(self, pkg: DocxPackage):
        self.pkg = pkg
        self.assets: dict[str, AssetBlob] = {}
        self.ole_blobs: dict[str, bytes] = {}
        self.issues: list[str] = []

    def asset_for_rel(self, rid: str | None, kind: str) -> AssetBlob | None:
        rel = self.pkg.rels.get(rid) if rid else None
        if rel is None or rel.external:
            return None
        name = self.pkg.part_name(rel)
        if not self.pkg.exists(name):
            self.issues.append(f"missing_part:{name}")
            return None
        data = self.pkg.read(name)
        sha = hashlib.sha256(data).hexdigest()
        ext = posixpath.splitext(name)[1].lstrip(".").lower()
        blob = self.assets.get(sha) or AssetBlob(
            sha, kind, MIME.get(ext, "application/octet-stream"), posixpath.basename(name), data
        )
        self.assets[sha] = blob
        return blob


# ------------------------------------------------------------------ nội dung inline


def _add_text(out: list[dict], text: str, marks: dict) -> None:
    if not text:
        return
    if out and out[-1]["type"] == "text" and out[-1].get("marks", {}) == marks:
        out[-1]["text"] += text
    else:
        node = {"type": "text", "text": text}
        if marks:
            node["marks"] = dict(marks)
        out.append(node)


def _object_node(obj, ctx: _Ctx) -> dict:
    ole = obj.find(q("o", "OLEObject"))
    progid = ole.get("ProgID", "") if ole is not None else ""
    ole_blob = ctx.asset_for_rel(ole.get(q("r", "id")) if ole is not None else None, "ole")
    img = obj.find(".//" + q("v", "imagedata"))
    prev = ctx.asset_for_rel(img.get(q("r", "id")) if img is not None else None, "preview")
    if ole_blob is not None and progid.startswith(("Equation", "DSMT")):
        ctx.ole_blobs[ole_blob.sha256] = ole_blob.data
        return {
            "type": "math", "source": None, "latex": None, "status": "pending",
            "ole": ole_blob.sha256, "preview": prev.sha256 if prev else None, "progid": progid,
        }
    ctx.issues.append(f"unsupported_ole:{progid or '?'}")
    return {
        "type": "object", "progid": progid, "status": "needs_review",
        "ole": ole_blob.sha256 if ole_blob else None, "preview": prev.sha256 if prev else None,
    }


def _drawing_node(drw, ctx: _Ctx) -> dict:
    blip = drw.find(".//" + q("a", "blip"))
    ext = drw.find(".//" + q("wp", "extent"))
    docpr = drw.find(".//" + q("wp", "docPr"))
    if blip is None:
        ctx.issues.append("unsupported_drawing_without_picture")
        return {"type": "unsupported", "what": "drawing", "status": "needs_review"}
    blob = ctx.asset_for_rel(blip.get(q("r", "embed")), "image")
    if blob is None:
        return {"type": "unsupported", "what": "image_missing", "status": "needs_review"}
    node = {"type": "image", "asset": blob.sha256}
    if ext is not None:
        node["width_emu"], node["height_emu"] = int(ext.get("cx", 0)), int(ext.get("cy", 0))
    if docpr is not None and docpr.get("descr"):
        node["alt"] = docpr.get("descr")
    node["floating"] = drw.find(q("wp", "anchor")) is not None
    src = blip.getparent().find(q("a", "srcRect")) if blip.getparent() is not None else None
    if src is not None:  # Word cắt bớt ảnh (đơn vị 1/1000 %): phải giữ để hiển thị giống Word
        crop = {k: int(src.get(k)) for k in "ltrb" if src.get(k) and int(src.get(k)) != 0}
        if crop:
            node["crop"] = crop
    return node


def _run_children(container, ctx: _Ctx, marks: dict, out: list[dict]) -> None:
    for ch in container:
        if not isinstance(ch.tag, str):
            continue
        ns, name = _ns(ch), _local(ch)
        if ns == NS["w"]:
            if name == "t":
                _add_text(out, ch.text or "", marks)
            elif name in ("tab", "noBreakHyphen"):
                _add_text(out, "\t" if name == "tab" else "-", marks)
            elif name in ("br", "cr"):
                if ch.get(q("w", "type")) != "page":
                    out.append({"type": "br"})
            elif name == "object":
                out.append(_object_node(ch, ctx))
            elif name == "drawing":
                out.append(_drawing_node(ch, ctx))
            elif name == "pict":
                img = ch.find(".//" + q("v", "imagedata"))
                if ch.find(".//" + q("v", "textbox")) is not None or img is None:
                    ctx.issues.append("unsupported_pict")
                    out.append({"type": "unsupported", "what": "pict", "status": "needs_review"})
                else:
                    blob = ctx.asset_for_rel(img.get(q("r", "id")), "image")
                    if blob:
                        out.append({"type": "image", "asset": blob.sha256, "floating": False})
        elif ns == NS["mc"] and name == "AlternateContent":
            choice = ch.find(q("mc", "Choice"))
            _run_children(choice if choice is not None else ch.find(q("mc", "Fallback")), ctx, marks, out)


def _inlines(p, ctx: _Ctx) -> list[dict]:
    out: list[dict] = []

    def walk(el, inherited: dict):
        for ch in el:
            if not isinstance(ch.tag, str):
                continue
            ns, name = _ns(ch), _local(ch)
            if ns == NS["w"] and name == "r":
                _run_children(ch, ctx, {**inherited, **_marks(ch.find(q("w", "rPr")))}, out)
            elif ns == NS["w"] and name in ("hyperlink", "smartTag", "ins", "sdt", "sdtContent", "fldSimple", "customXml"):
                walk(ch, inherited)
            elif ns == NS["m"] and name in ("oMath", "oMathPara"):
                ctx.issues.append("omml_not_supported_yet")
                out.append({
                    "type": "math", "source": "omml", "latex": None, "status": "needs_review",
                    "omml": etree.tostring(ch, encoding="unicode"),
                })
            elif ns == NS["mc"] and name == "AlternateContent":
                choice = ch.find(q("mc", "Choice"))
                walk(choice if choice is not None else ch, inherited)

    walk(p, {})
    return out


def _block(el, ctx: _Ctx) -> dict | None:
    name = _local(el)
    if name == "p":
        return {"type": "paragraph", "inlines": _inlines(el, ctx)}
    if name == "tbl":
        rows = []
        for tr in el.findall(q("w", "tr")):
            cells = []
            for tc in tr.findall(q("w", "tc")):
                cells.append([b for b in (_block(c, ctx) for c in tc if _local(c) in ("p", "tbl")) if b])
            rows.append(cells)
        return {"type": "table", "rows": rows}
    return None


def _strip_question_marker(blocks: list[dict]) -> None:
    """Bỏ nhãn gõ tay "Câu N." đầu câu khỏi nội dung (số câu là metadata, không phải nội dung)."""
    if not blocks or blocks[0]["type"] != "paragraph":
        return
    inl = blocks[0]["inlines"]
    while inl and inl[0]["type"] == "text" and not inl[0]["text"].strip():
        inl.pop(0)
    if inl and inl[0]["type"] == "text":
        inl[0]["text"] = QUESTION_LABEL_RE.sub("", inl[0]["text"], count=1)
        if not inl[0]["text"].strip():
            inl.pop(0)


# ------------------------------------------------------------------ tách phương án A-D


def _has_content(inl: list[dict]) -> bool:
    return any(n["type"] != "text" or n["text"].strip() for n in inl)


def _split_options(inlines: list[dict]) -> list[dict] | None:
    """Tách A./B./C./D. theo đúng thứ tự; trả None nếu không đủ chuỗi A→D."""
    opts: list[dict] = []
    expected = "A"
    for node in inlines:
        if node["type"] == "text":
            text, pos = node["text"], 0
            for m in OPTION_RE.finditer(text):
                if m.group(1) != expected:
                    continue
                before = text[pos : m.start()]
                if opts:
                    _add_text(opts[-1]["inlines"], before, node.get("marks", {}))
                opts.append({"label": expected, "inlines": []})
                expected = chr(ord(expected) + 1)
                pos = m.end()
            if opts:
                _add_text(opts[-1]["inlines"], text[pos:], node.get("marks", {}))
        elif opts:
            opts[-1]["inlines"].append(node)
    if [o["label"] for o in opts] != list("ABCD"):
        return None
    for o in opts:
        inl = o["inlines"]
        while inl and inl[-1]["type"] == "text" and not inl[-1]["text"].strip():
            inl.pop()
        if inl and inl[-1]["type"] == "text":
            inl[-1]["text"] = inl[-1]["text"].rstrip()
        if inl and inl[0]["type"] == "text":
            inl[0]["text"] = inl[0]["text"].lstrip()
    return opts


def _build_content(blocks: list[dict]) -> tuple[dict, str]:
    """Từ các khối của một câu → {"stem": [...], "options": [...]|None} và qtype."""
    opt_start = None
    for i, b in enumerate(blocks):
        if i > 0 and b["type"] == "paragraph":
            first = next((n for n in b["inlines"] if n["type"] != "text" or n["text"].strip()), None)
            if first and first["type"] == "text" and re.match(r"\s*A\.(\s|$)", first["text"]):
                opt_start = i
                break
    if opt_start is not None:
        flat: list[dict] = []
        for j in range(opt_start, len(blocks)):
            b = blocks[j]
            if b["type"] == "paragraph":
                if flat:
                    flat.append({"type": "text", "text": "  "})
                flat.extend(b["inlines"])
            options = _split_options(flat)
            if options:  # đủ A→D: dừng, phần còn lại (ghi chú, hình…) là "after"
                content = {"stem": blocks[:opt_start], "options": options}
                if blocks[j + 1 :]:
                    content["after"] = blocks[j + 1 :]
                return content, "multiple_choice"
    return {"stem": blocks, "options": None}, "free_response"


def _text_of(els) -> str:
    return " ".join(para_text(e) for e in els if _local(e) == "p")


def _relocate_figures(groups: list[dict]) -> None:
    """Hình nổi bị neo nhầm vào câu liền trước (thường do giáo viên kéo thả) → chuyển sang câu hình học kế tiếp.

    Chỉ chuyển khi câu hiện tại KHÔNG nói về hình học, câu kế tiếp (cùng đề, cùng phần) CÓ. Câu được gắn cờ để xem lại.
    """
    for g, nxt in zip(groups, groups[1:]):
        if (g["exam"], g["section"]) != (nxt["exam"], nxt["section"]):
            continue
        if GEOM_RE.search(_text_of(g["els"])) or not GEOM_RE.search(_text_of(nxt["els"])):
            continue
        target = next((e for e in nxt["els"] if _local(e) == "p"), None)
        if target is None:
            continue
        runs = []
        for e in g["els"]:
            for a in e.iter(q("wp", "anchor")):
                r = a
                while r is not None and r.tag != q("w", "r"):
                    r = r.getparent()
                if r is not None and r not in runs:
                    runs.append(r)
        for i, r in enumerate(runs):
            r.getparent().remove(r)
            ppr = target.find(q("w", "pPr"))
            target.insert((list(target).index(ppr) + 1 if ppr is not None else 0) + i, r)
        if runs:
            nxt["moved"] = True


# ------------------------------------------------------------------ điểm vào chính


def parse_document(pkg: DocxPackage) -> ParsedDocument:
    ctx = _Ctx(pkg)
    numbering = pkg.numbering
    body = pkg.document.find(q("w", "body"))
    questions: list[ParsedQuestion] = []

    exam: str | None = None
    section: str | None = None
    counter = 0
    current: list = []  # các phần tử XML của câu đang gom

    groups: list[dict] = []

    def flush():
        nonlocal current, counter
        if not current:
            return
        els = list(current)
        current = []
        while len(els) > 1 and _local(els[-1]) == "p" and not para_text(els[-1]).strip() and \
                els[-1].find(".//" + q("w", "drawing")) is None and els[-1].find(".//" + q("w", "object")) is None:
            els.pop()
        counter += 1
        groups.append({"exam": exam, "section": section, "number": counter, "els": els})

    def build(g: dict) -> ParsedQuestion:
        els = g["els"]
        blocks = [b for b in (_block(e, ctx) for e in els) if b]
        _strip_question_marker(blocks)
        content, qtype = _build_content(blocks)
        pm = POINTS_RE.search(para_text(els[0]))
        rel_ids = {
            v for e in els for node in e.iter() if isinstance(node.tag, str)
            for k, v in node.attrib.items() if k.startswith("{%s}" % NS["r"])
        }
        return ParsedQuestion(
            exam=g["exam"], section=g["section"], number=g["number"], qtype=qtype,
            points=float(pm.group(1).replace(",", ".")) if pm else None,
            content=content,
            raw_ooxml="".join(etree.tostring(e, encoding="unicode") for e in els),
            rel_ids=rel_ids, issues=[MOVED_FIGURE] if g.get("moved") else [],
        )

    for el in body:
        name = _local(el)
        if name not in ("p", "tbl"):
            continue
        text = para_text(el) if name == "p" else ""
        num = el.find(q("w", "pPr") + "/" + q("w", "numPr")) if name == "p" else None
        is_list_q = False
        if num is not None:
            nid = num.find(q("w", "numId"))
            lv = num.find(q("w", "ilvl"))
            is_list_q = nid is not None and numbering.is_ordered_top_level(
                nid.get(q("w", "val")), lv.get(q("w", "val")) if lv is not None else "0"
            )
        if name == "p" and num is None and EXAM_RE.match(text):
            flush()
            exam, section, counter = text.strip(), None, 0
        elif name == "p" and num is None and SECTION_RE.match(text) and len(text) < 160:
            flush()
            section = text.strip()
        elif name == "p" and (is_list_q or QUESTION_RE.match(text)):
            flush()
            current = [el]
        elif current:
            current.append(el)
    flush()
    _relocate_figures(groups)
    questions.extend(build(g) for g in groups)

    return ParsedDocument(questions, ctx.assets, ctx.ole_blobs, ctx.issues)
