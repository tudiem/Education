"""Nhập một file Word vào kho: cắt câu → chuyển công thức → kiểm tra → lưu DB + tài sản."""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

from sqlalchemy import select

from . import latexcheck, mathtype
from .assets import AssetStore
from .db import Asset, Question, SourceFile, Topic, make_code
from .package import DocxPackage
from .parse import ParsedDocument, parse_document


class AlreadyImported(Exception):
    def __init__(self, source_file_id: int):
        super().__init__(f"File này đã được nhập vào chủ đề (source_file #{source_file_id})")
        self.source_file_id = source_file_id


@dataclass
class ImportReport:
    filename: str
    topic: str
    exams: list[str] = field(default_factory=list)
    questions: int = 0
    multiple_choice: int = 0
    free_response: int = 0
    formulas_total: int = 0  # số lần xuất hiện
    formulas_unique: int = 0
    formulas_by_source: dict[str, int] = field(default_factory=dict)
    formulas_needs_review: int = 0
    questions_needs_review: int = 0
    duplicates: int = 0
    images: int = 0
    katex_validation: str = "skipped"
    fixes: dict[str, int] = field(default_factory=dict)
    issues: list[str] = field(default_factory=list)
    codes: list[str] = field(default_factory=list)


def _walk_inlines(content: dict):
    """Duyệt mọi nút inline (cả trong bảng) của một câu."""
    def blocks(bl):
        for b in bl:
            if b["type"] == "paragraph":
                yield from b["inlines"]
            elif b["type"] == "table":
                for row in b["rows"]:
                    for cell in row:
                        yield from blocks(cell)
    yield from blocks(content["stem"])
    yield from blocks(content.get("after", []))
    for o in content.get("options") or []:
        yield from o["inlines"]


def _resolve_formulas(doc: ParsedDocument, use_mathml: bool, report: ImportReport) -> None:
    formulas = mathtype.convert_many(doc.ole_blobs, use_mathml=use_mathml)

    latexes = sorted({f.latex for f in formulas.values() if f.latex})
    errs = latexcheck.validate(latexes)
    if errs is not None:
        report.katex_validation = "passed" if not any(errs) else "failed"
        bad = {lx: e for lx, e in zip(latexes, errs) if e}
        for f in formulas.values():
            if f.latex in bad:
                f.status = "needs_review"
                f.notes.append(f"katex:{bad[f.latex]}")

    for q in doc.questions:
        for node in _walk_inlines(q.content):
            if node["type"] != "math" or node.get("source") == "omml":
                continue
            f = formulas[node["ole"]]
            node.update(latex=f.latex, source=f.source, status=f.status)
            if f.notes:
                node["notes"] = f.notes
    report.formulas_unique = len(formulas)
    for f in formulas.values():
        report.formulas_by_source[f.source] = report.formulas_by_source.get(f.source, 0) + 1
        for n in f.notes:
            if n.startswith("fix:"):
                report.fixes[n] = report.fixes.get(n, 0) + 1


def import_docx(
    path: str | Path, topic_name: str, Session, store: AssetStore, *, use_mathml: bool = True
) -> ImportReport:
    path = Path(path)
    raw = path.read_bytes()
    file_sha = hashlib.sha256(raw).hexdigest()
    pkg = DocxPackage(raw)
    doc = parse_document(pkg)
    report = ImportReport(filename=path.name, topic=topic_name, issues=list(dict.fromkeys(doc.issues)))
    _resolve_formulas(doc, use_mathml, report)

    with Session() as s, s.begin():
        topic = s.scalar(select(Topic).where(Topic.name == topic_name))
        if topic is None:
            topic = Topic(name=topic_name)
            s.add(topic)
            s.flush()
        existing = s.scalar(select(SourceFile).where(SourceFile.topic_id == topic.id, SourceFile.sha256 == file_sha))
        if existing:
            raise AlreadyImported(existing.id)
        sf = SourceFile(topic_id=topic.id, filename=path.name, sha256=file_sha)
        s.add(sf)
        s.flush()

        store.put(file_sha, raw)  # giữ file nguồn: khuôn style/đánh số khi xuất .docx
        if s.get(Asset, file_sha) is None:
            s.add(Asset(sha256=file_sha, kind="source", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", size=len(raw), name=path.name))
        for sha, blob in doc.assets.items():
            store.put(sha, blob.data)
            if s.get(Asset, sha) is None:
                s.add(Asset(sha256=sha, kind=blob.kind, mime=blob.mime, size=len(blob.data), name=blob.name))
        report.images = sum(1 for b in doc.assets.values() if b.kind == "image")

        seen: dict[str, str] = {
            h: c for h, c in s.execute(
                select(Question.content_hash, Question.code).where(Question.topic_id == topic.id)
            )
        }
        exams: list[str] = []
        for pq in doc.questions:
            nodes = list(_walk_inlines(pq.content))
            math = [n for n in nodes if n["type"] == "math"]
            report.formulas_total += len(math)
            flagged = [n for n in nodes if n.get("status") == "needs_review"]
            issues = [
                f"{n['type']}:{','.join(n.get('notes', [])) or n.get('what') or n.get('progid') or 'review'}"
                for n in flagged
            ]
            issues += pq.issues
            chash = hashlib.sha256(
                json.dumps(pq.content, sort_keys=True, ensure_ascii=False).encode()
            ).hexdigest()

            rels = {}
            for rid in sorted(pq.rel_ids):
                rel = pkg.rels.get(rid)
                if rel is None:
                    continue
                entry = {"type": rel.type.rsplit("/", 1)[-1], "target": rel.target, "external": rel.external}
                if not rel.external:
                    name = pkg.part_name(rel)
                    if name.startswith(("word/media/", "word/embeddings/")) and pkg.exists(name):
                        data = pkg.read(name)
                        sha = hashlib.sha256(data).hexdigest()
                        store.put(sha, data)
                        if s.get(Asset, sha) is None and sha not in doc.assets:
                            s.add(Asset(sha256=sha, kind="ole" if "embeddings" in name else "image",
                                        mime="application/octet-stream", size=len(data), name=name.rsplit("/", 1)[-1]))
                        entry["asset"] = sha
                rels[rid] = entry
            s.flush()

            q = Question(
                topic_id=topic.id, source_file_id=sf.id, exam=pq.exam, section=pq.section, number=pq.number,
                qtype=pq.qtype, points=pq.points, content=pq.content, raw_ooxml=pq.raw_ooxml, raw_rels=rels,
                status="needs_review" if (flagged or pq.issues) else "ok", issues=issues, content_hash=chash,
                duplicate_of=seen.get(chash),
            )
            s.add(q)
            s.flush()
            q.code = make_code(q.id)
            seen.setdefault(chash, q.code)
            report.codes.append(q.code)
            report.questions += 1
            report.multiple_choice += pq.qtype == "multiple_choice"
            report.free_response += pq.qtype == "free_response"
            report.duplicates += q.duplicate_of is not None
            report.questions_needs_review += bool(flagged or pq.issues)
            report.formulas_needs_review += sum(1 for n in math if n.get("status") == "needs_review")
            if pq.exam and pq.exam not in exams:
                exams.append(pq.exam)
        report.exams = exams
        rep = asdict(report)
        rep.pop("codes")
        sf.report = rep
    return report
