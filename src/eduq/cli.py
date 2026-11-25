"""Dòng lệnh: eduq import | list | show | preview."""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from html import escape
from pathlib import Path

from sqlalchemy import select

from . import export as exp
from . import render
from .assets import AssetStore
from .db import Question, Topic, open_db
from .importer import AlreadyImported, import_docx
from .mathtype import TOOLS_DIR
from .package import DocxError

PREVIEW_CSS = (
    ".q{border:1px solid #ccc;border-radius:6px;padding:10px 14px;margin:12px 0}"
    ".q.needs_review{border-color:#d33;background:#fff6f6}.q:after{content:'';display:block;clear:both}"
    "body{max-width:900px;margin:24px auto;padding:0 16px}"
)


def _env(args):
    data = Path(args.data)
    data.mkdir(parents=True, exist_ok=True)
    return open_db(f"sqlite:///{data / 'eduq.db'}"), AssetStore(data / "assets")


def cmd_import(args) -> int:
    Session, store = _env(args)
    try:
        r = import_docx(args.file, args.topic, Session, store, use_mathml=not args.no_mathml)
    except (DocxError, AlreadyImported) as e:
        print(f"Lỗi: {e}", file=sys.stderr)
        return 1
    print(f"Đã nhập {r.questions} câu ({r.codes[0]}…{r.codes[-1]}) vào chủ đề “{r.topic}”")
    print(f"  Đề: {len(r.exams)} | trắc nghiệm: {r.multiple_choice} | tự luận: {r.free_response} | hình: {r.images}")
    print(f"  Công thức: {r.formulas_total} (duy nhất {r.formulas_unique}); nguồn: {r.formulas_by_source}")
    print(f"  Tự sửa ký hiệu: {r.fixes or 'không'} | KaTeX: {r.katex_validation}")
    print(f"  Cần review: {r.questions_needs_review} câu / {r.formulas_needs_review} công thức | trùng nội dung: {r.duplicates}")
    for i in r.issues:
        print(f"  ! {i}")
    return 0


def cmd_list(args) -> int:
    Session, _ = _env(args)
    with Session() as s:
        stmt = select(Question).order_by(Question.id)
        if args.status:
            stmt = stmt.where(Question.status == args.status)
        for q in s.scalars(stmt.limit(args.limit)):
            print(f"{q.code}  {q.exam or '-':10} #{q.number:<2} {q.qtype:15} {q.status}")
    return 0


def cmd_show(args) -> int:
    Session, _ = _env(args)
    with Session() as s:
        q = s.scalar(select(Question).where(Question.code == args.code))
        if not q:
            print("Không tìm thấy", args.code, file=sys.stderr)
            return 1
        print(json.dumps(
            {"code": q.code, "exam": q.exam, "section": q.section, "number": q.number, "type": q.qtype,
             "points": q.points, "status": q.status, "issues": q.issues, "content": q.content},
            ensure_ascii=False, indent=2,
        ))
    return 0


def cmd_preview(args) -> int:
    Session, store = _env(args)
    out = Path(args.out)
    (out / "assets").mkdir(parents=True, exist_ok=True)
    kdist = TOOLS_DIR / "node_modules" / "katex" / "dist"
    if not kdist.is_dir():
        print("Thiếu KaTeX: chạy `npm install` trong thư mục tools/", file=sys.stderr)
        return 1
    shutil.copytree(kdist, out / "katex", dirs_exist_ok=True)
    used: set[str] = set()

    def asset_url(sha: str) -> str:
        used.add(sha)
        return f"assets/{sha}"

    parts = []
    with Session() as s:
        stmt = select(Question).order_by(Question.id)
        if args.exam:
            stmt = stmt.where(Question.exam == args.exam)
        for q in s.scalars(stmt):
            parts.append(
                f'<div class="q {q.status}"><div class="hd"><b>{q.code}</b> · {escape(q.exam or "")} · '
                f'{escape(q.section or "")} · câu {q.number} · {q.qtype}'
                f'{" · ⚠ " + escape("; ".join(q.issues)) if q.issues else ""}</div>'
                f"{render.question_html(q.content, asset_url)}</div>"
            )
    for sha in used:
        shutil.copyfile(store.path(sha), out / "assets" / sha)
    (out / "index.html").write_text(
        render.page_html("\n".join(parts), title="Xem trước kho bài tập", extra_css=PREVIEW_CSS), encoding="utf-8"
    )
    print(f"Đã tạo {out / 'index.html'} ({len(parts)} câu)")
    return 0


def cmd_export(args) -> int:
    Session, store = _env(args)
    try:
        codes = exp.parse_codes(args.codes) if args.codes else None
        if not (codes or args.exam or args.topic):
            raise exp.ExportError("Cần chọn câu hỏi: --codes, --exam hoặc --topic")
        out = Path(args.out)
        fmts = ["docx", "pdf"] if args.format == "both" else [args.format]
        with Session() as s:
            qs = exp.select_questions(s, codes=codes, exam=args.exam, topic=args.topic)
            title = args.title if args.title is not None else args.exam
            print(f"Chọn {len(qs)} câu: {qs[0].code}…{qs[-1].code}")
            results = []
            if "docx" in fmts:
                tpl, w = exp.template_for(s, store, qs)
                r = exp.export_docx(qs, tpl, store, out / f"{args.name}.docx", title, args.paper)
                r.warnings = w + r.warnings
                results.append(r)
            if "pdf" in fmts:
                results.append(exp.export_pdf(qs, s, store, out / f"{args.name}.pdf", title, args.pdf_engine, args.paper))
    except exp.ExportError as e:
        print(f"Lỗi: {e}", file=sys.stderr)
        return 1
    for r in results:
        print(f"  → {r.path} ({r.path.stat().st_size // 1024} KB)")
        for w in r.warnings:
            print(f"    ! {w}")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="eduq", description="Kho bài tập: nhập từ Word (MathType)")
    ap.add_argument("--data", default="data", help="thư mục dữ liệu (DB + tài sản), mặc định ./data")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("import", help="nhập file .docx vào một chủ đề")
    p.add_argument("file")
    p.add_argument("--topic", required=True)
    p.add_argument("--no-mathml", action="store_true", help="không dùng gem Ruby/pandoc (chỉ đường TeX nhúng)")
    p.set_defaults(fn=cmd_import)
    p = sub.add_parser("list", help="liệt kê câu hỏi")
    p.add_argument("--status", choices=["ok", "needs_review"])
    p.add_argument("--limit", type=int, default=50)
    p.set_defaults(fn=cmd_list)
    p = sub.add_parser("show", help="xem một câu theo mã, vd. C0000001")
    p.add_argument("code")
    p.set_defaults(fn=cmd_show)
    p = sub.add_parser("preview", help="xuất trang HTML xem trước (KaTeX)")
    p.add_argument("--out", default="preview")
    p.add_argument("--exam", help='lọc theo đề, vd. "ĐỀ SỐ 1"')
    p.set_defaults(fn=cmd_preview)
    p = sub.add_parser("export", help="xuất câu hỏi ra .docx / PDF")
    p.add_argument("--codes", help="mã hoặc khoảng, vd. C0000001,C0000005-C0000010 (giữ đúng thứ tự)")
    p.add_argument("--exam", help='lấy cả đề, vd. "ĐỀ SỐ 1"')
    p.add_argument("--topic", help="lấy cả chủ đề")
    p.add_argument("--title", help="tiêu đề đầu đề (mặc định: tên đề nếu có --exam)")
    p.add_argument("--format", choices=["docx", "pdf", "both"], default="both")
    p.add_argument("--pdf-engine", choices=["auto", "chrome", "libreoffice"], default="auto")
    p.add_argument("--paper", choices=["a4", "source"], default="a4", help="khổ giấy: A4 (mặc định) hoặc giữ như file nguồn")
    p.add_argument("--out", default="export")
    p.add_argument("--name", default="de-thi", help="tên file (không đuôi)")
    p.set_defaults(fn=cmd_export)
    args = ap.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    raise SystemExit(main())
