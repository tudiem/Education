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


def question_html(content: dict, asset_url, label: str | None = None) -> str:
    """HTML một câu. `label` (vd. "Câu 3:") được chèn đầu đoạn đầu tiên."""
    out = "".join(_block(b, asset_url) for b in content["stem"])
    if label:
        tag = f'<span class="no">{escape(label)}</span> '
        out = out.replace("<p>", "<p>" + tag, 1) if out.startswith("<p>") else f"<p>{tag}</p>" + out
    if content.get("options"):
        out += '<ol class="opts">' + "".join(
            f'<li><b>{o["label"]}.</b> ' + "".join(_inline(n, asset_url) for n in o["inlines"]) + "</li>"
            for o in content["options"]
        ) + "</ol>"
        out += "".join(_block(b, asset_url) for b in content.get("after", []))
    return out


PAGE_CSS = """
body{font:16px/1.6 "Times New Roman","Liberation Serif",serif;color:#111}
.hd{font:12px sans-serif;color:#555;margin-bottom:4px}.hd b{color:#06c}
.opts{list-style:none;padding:0;display:flex;gap:28px;flex-wrap:wrap}.opts li{min-width:120px}
img.pic{display:block;margin:6px 0}img.pic.float{float:right;margin:0 0 8px 16px;max-width:45%}
img.fallback{height:1.4em;vertical-align:middle}
span.review{background:#fdd}span.no{font-weight:bold;color:#0070c0}
p{margin:4px 0}table{border-collapse:collapse}td{padding:2px 12px}
"""

KATEX_RENDER_JS = (
    "document.querySelectorAll('span.math').forEach(function(e){"
    "try{katex.render(e.dataset.latex,e,{throwOnError:false})}catch(x){e.textContent=e.dataset.latex}})"
)


def page_html(body: str, *, title: str = "Kho bài tập", extra_css: str = "") -> str:
    """Trang HTML hoàn chỉnh; cần thư mục `katex/` (css + js + fonts) nằm cạnh file."""
    return (
        f'<!doctype html><html lang="vi"><meta charset="utf-8"><title>{escape(title)}</title>'
        f'<link rel="stylesheet" href="katex/katex.min.css"><style>{PAGE_CSS}{extra_css}</style><body>'
        f'{body}<script src="katex/katex.min.js"></script><script>{KATEX_RENDER_JS}</script></body></html>'
    )
