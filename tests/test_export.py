import os
import shutil
import zipfile
from pathlib import Path

import pytest
from conftest import PNG, make_docx, p, p_image, p_split
from lxml import etree
from sqlalchemy import select

from eduq import export as exp
from eduq.assets import AssetStore
from eduq.db import Question, open_db
from eduq.importer import import_docx
from eduq.package import NS, q

NUMBERING = (
    '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="Câu %1:"/><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="10"><w:abstractNumId w:val="3"/></w:num>'
)


@pytest.fixture()
def env(tmp_path):
    return open_db(f"sqlite:///{tmp_path}/t.db"), AssetStore(tmp_path / "assets"), tmp_path


def add(tmp, name, body, numbering=None, images=None):
    f = tmp / name
    f.write_bytes(make_docx(body, numbering, images))
    return f


def do_export(Session, store, codes, out, **kw):
    with Session() as s:
        qs = exp.select_questions(s, codes=codes)
        tpl, w = exp.template_for(s, store, qs)
        res = exp.export_docx(qs, tpl, store, out, **kw)
        res.warnings = w + res.warnings
        return res


def doc_xml(path):
    return etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))


# ---------------------------------------------------------------- chọn câu


def test_parse_codes():
    assert exp.parse_codes("C0000003, C0000001-C0000002,c0000003") == ["C0000003", "C0000001", "C0000002"]
    for bad in ("C1", "X0000001", "C0000005-C0000001"):
        with pytest.raises(exp.ExportError):
            exp.parse_codes(bad)


def test_select_unknown_code_errors(env):
    Session, _, tmp = env
    with Session() as s, pytest.raises(exp.ExportError):
        exp.select_questions(s, codes=["C0000999"])


# ---------------------------------------------------------------- .docx


def test_export_order_numbering_and_roundtrip(env):
    Session, store, tmp = env
    f = add(tmp, "a.docx", p("Tính 1+1?", num_id="10") + p("Tính 2+2?", num_id="10") + p("Tính 3+3?", num_id="10"), NUMBERING)
    import_docx(f, "Toán", Session, store)
    out = tmp / "out.docx"
    res = do_export(Session, store, ["C0000003", "C0000001"], out, title="Đề kiểm tra")
    assert res.warnings == [] and exp.check_docx(out) == []

    body = doc_xml(out).find(q("w", "body"))
    paras = [x for x in body if x.tag == q("w", "p")]
    texts = ["".join(t.text for t in x.iter(q("w", "t"))) for x in paras]
    assert texts == ["Đề kiểm tra", "Tính 3+3?", "Tính 1+1?"]  # đúng thứ tự yêu cầu
    num_ids = {x.find(".//" + q("w", "numId")).get(q("w", "val")) for x in paras[1:]}
    assert len(num_ids) == 1 and num_ids != {"10"}  # danh sách "Câu n:" mới, không đụng numId cũ
    assert body[-1].tag == q("w", "sectPr") and not body.findall(".//" + q("w", "headerReference"))

    # numbering.xml: abstractNum luôn đứng trước num (đúng schema)
    nx = etree.fromstring(zipfile.ZipFile(out).read("word/numbering.xml"))
    tags = [etree.QName(c).localname for c in nx]
    assert tags.index("num") > max(i for i, t in enumerate(tags) if t == "abstractNum")

    # vòng tròn: nhập lại file vừa xuất → nội dung từng câu giống hệt
    Session2 = open_db(f"sqlite:///{tmp}/rt.db")
    import_docx(out, "rt", Session2, AssetStore(tmp / "assets2"))
    with Session() as a, Session2() as b:
        orig = {x.code: x.content_hash for x in a.scalars(select(Question))}
        got = [x.content_hash for x in b.scalars(select(Question).order_by(Question.id))]
    assert got == [orig["C0000003"], orig["C0000001"]]


def test_typed_label_stripped_even_when_split_across_runs(env):
    Session, store, tmp = env
    f = add(tmp, "t.docx", p_split("Câu ", "3. Nội dung") + p("Bài 4: Câu kế"))
    import_docx(f, "T", Session, store)
    out = tmp / "o.docx"
    do_export(Session, store, ["C0000001", "C0000002"], out)
    paras = doc_xml(out).find(q("w", "body")).findall(q("w", "p"))
    assert ["".join(t.text or "" for t in x.iter(q("w", "t"))).strip() for x in paras] == ["Nội dung", "Câu kế"]
    assert all(x.find(".//" + q("w", "numPr")) is not None for x in paras)


def test_images_copied_and_rels_remapped(env):
    Session, store, tmp = env
    f = add(tmp, "i.docx", p_image("Câu có hình", "rId7", num_id="10") + p_image("Câu hai", "rId7", num_id="10"),
            NUMBERING, images={"rId7": PNG})
    import_docx(f, "Hình", Session, store)
    out = tmp / "o.docx"
    res = do_export(Session, store, ["C0000001", "C0000002"], out)
    assert res.warnings == [] and exp.check_docx(out) == []
    z = zipfile.ZipFile(out)
    media = [n for n in z.namelist() if n.startswith("word/media/")]
    assert len(media) == 1 and z.read(media[0]) == PNG  # cùng ảnh → 1 part
    ids = [d.get("id") for d in doc_xml(out).iter("{%s}docPr" % NS["wp"])]
    assert ids == ["1", "2"]  # id hình được đánh lại duy nhất
    Session2 = open_db(f"sqlite:///{tmp}/rt.db")
    r = import_docx(out, "rt", Session2, AssetStore(tmp / "a2"))
    assert r.images == 1 and r.questions == 2


def test_multiple_sources_unified_numbering(env):
    Session, store, tmp = env
    import_docx(add(tmp, "a.docx", p("Từ file A", num_id="10"), NUMBERING), "T", Session, store)
    import_docx(add(tmp, "b.docx", p("Câu 1. Từ file B")), "T", Session, store)
    out = tmp / "o.docx"
    res = do_export(Session, store, ["C0000002", "C0000001"], out)
    assert any("2 file nguồn" in w for w in res.warnings)
    assert exp.check_docx(out) == []


def test_source_missing_falls_back_to_default_style(env):
    Session, store, tmp = env
    f = add(tmp, "a.docx", p("Câu 1. X"))
    import_docx(f, "T", Session, store)
    for fp in (tmp / "assets").rglob("*"):  # mô phỏng dữ liệu nhập từ phiên bản cũ (chưa lưu file nguồn)
        if fp.is_file() and fp.read_bytes()[:2] == b"PK":
            fp.unlink()
    out = tmp / "o.docx"
    res = do_export(Session, store, ["C0000001"], out)
    assert any("mặc định" in w for w in res.warnings) and exp.check_docx(out) == []


def test_paper_a4_vs_source(env):
    Session, store, tmp = env
    import_docx(add(tmp, "a.docx", p("Câu 1. X")), "T", Session, store)
    for paper, w in (("a4", "11906"), ("source", None)):
        out = tmp / f"{paper}.docx"
        do_export(Session, store, ["C0000001"], out, paper=paper)
        pg = doc_xml(out).find(".//" + q("w", "pgSz"))
        if w:
            assert pg.get(q("w", "w")) == w
        else:
            assert pg is None or pg.get(q("w", "w")) is not None


def test_check_docx_detects_problems(env, tmp_path):
    Session, store, tmp = env
    import_docx(add(tmp, "a.docx", p("Câu 1. X")), "T", Session, store)
    good = tmp / "g.docx"
    do_export(Session, store, ["C0000001"], good)
    bad = tmp / "b.docx"
    with zipfile.ZipFile(good) as zi, zipfile.ZipFile(bad, "w") as zo:
        for n in zi.namelist():
            data = zi.read(n)
            if n == "word/document.xml":
                data = data.replace(b"<w:sectPr", b'<w:r><w:pict><v:imagedata xmlns:v="urn:schemas-microsoft-com:vml" r:id="rId99"/></w:pict></w:r><w:sectPr', 1)
            zo.writestr(n, data)
    assert any("rId99" in x for x in exp.check_docx(bad))


# ---------------------------------------------------------------- PDF


@pytest.mark.skipif("chrome" not in exp.pdf_engines(), reason="không có Chrome + KaTeX")
def test_pdf_chrome_produces_pdf(env):
    Session, store, tmp = env
    import_docx(add(tmp, "a.docx", p("Câu 1. Nội dung một") + p("Câu 2. Nội dung hai")), "T", Session, store)
    with Session() as s:
        qs = exp.select_questions(s, topic="T")
        res = exp.export_pdf(qs, s, store, tmp / "o.pdf", title="Đề thử", engine="chrome")
    assert (tmp / "o.pdf").read_bytes()[:4] == b"%PDF" and res.questions == 2


# ---------------------------------------------------------------- file thật

SAMPLE = os.environ.get("EDUQ_SAMPLE_DOCX")


@pytest.mark.skipif(not SAMPLE or not Path(SAMPLE).exists(), reason="đặt EDUQ_SAMPLE_DOCX để chạy với file thật")
def test_real_sample_export_all_roundtrip(env):
    Session, store, tmp = env
    import_docx(SAMPLE, "Toán 7", Session, store)
    out = tmp / "all.docx"
    with Session() as s:
        qs = exp.select_questions(s, topic="Toán 7")
        tpl, _ = exp.template_for(s, store, qs)
        res = exp.export_docx(qs, tpl, store, out, title="10 đề")
        orig = [x.content_hash for x in qs]
    assert res.warnings == [] and exp.check_docx(out) == []
    Session2 = open_db(f"sqlite:///{tmp}/rt.db")
    r = import_docx(out, "rt", Session2, AssetStore(tmp / "a2"))
    assert (r.questions, r.formulas_total, r.images, r.questions_needs_review) == (80, 397, 10, 0)  # file xuất: hình đã đúng câu
    with Session2() as b:
        assert [x.content_hash for x in b.scalars(select(Question).order_by(Question.id))] == orig
