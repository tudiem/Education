"""Hiển thị nội dung câu hỏi thành HTML (công thức để KaTeX dựng ở trình duyệt)."""
from __future__ import annotations

from html import escape

EMU_PER_PX = 9525


def _inline(n: dict, asset_url) -> str:
    t = n["type"]
    if t == "text":
        s = escape(n["text"]).replace("\t", "&emsp;")
        m = n.get("marks", {})
        for key, tag in (("b", "strong"), ("i", "em"), ("u", "u"), ("sup", "sup"), ("sub", "sub")):
            if m.get(key):
                s = f"<{tag}>{s}</{tag}>"
        return s
    if t == "br":
        return "<br>"
    if t == "math":
        cls = "math" + (" review" if n.get("status") != "ok" else "")
        if n.get("latex"):
            return f'<span class="{cls}" data-latex="{escape(n["latex"], quote=True)}"></span>'
        if n.get("preview"):  # chưa chuyển được: dùng ảnh gốc
            return f'<img class="fallback" src="{asset_url(n["preview"])}" alt="công thức (ảnh gốc)">'
        return '<span class="review">[công thức lỗi]</span>'
    if t == "image":
        w = n.get("width_emu")
        size = f' style="width:{w // EMU_PER_PX}px;max-width:100%"' if w else ""
        cls = "pic float" if n.get("floating") else "pic"
        return f'<img class="{cls}" src="{asset_url(n["asset"])}"{size} alt="{escape(n.get("alt", ""))}">'
    return '<span class="review">[nội dung chưa hỗ trợ]</span>'


def _block(b: dict, asset_url) -> str:
    if b["type"] == "paragraph":
        return "<p>" + "".join(_inline(n, asset_url) for n in b["inlines"]) + "</p>"
    if b["type"] == "table":
        rows = "".join(
            "<tr>" + "".join("<td>" + "".join(_block(x, asset_url) for x in cell) + "</td>" for cell in row) + "</tr>"
            for row in b["rows"]
        )
        return f"<table>{rows}</table>"
    return ""


def question_html(content: dict, asset_url) -> str:
    out = "".join(_block(b, asset_url) for b in content["stem"])
    if content.get("options"):
        out += '<ol class="opts">' + "".join(
            f'<li><b>{o["label"]}.</b> ' + "".join(_inline(n, asset_url) for n in o["inlines"]) + "</li>"
            for o in content["options"]
        ) + "</ol>"
        out += "".join(_block(b, asset_url) for b in content.get("after", []))
    return out
