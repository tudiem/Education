import os
import re
from pathlib import Path

import pytest
from conftest import make_docx, p
from sqlalchemy import select

from eduq.assets import AssetStore
from eduq.db import Question, open_db
from eduq.importer import AlreadyImported, import_docx


@pytest.fixture()
def env(tmp_path):
    return open_db(f"sqlite:///{tmp_path}/t.db"), AssetStore(tmp_path / "assets"), tmp_path


def write(tmp, name, body):
    f = tmp / name
    f.write_bytes(make_docx(body))
    return f


def test_codes_sequential_and_reimport_blocked(env):
    Session, store, tmp = env
    f = write(tmp, "a.docx", p("ĐỀ SỐ 1") + p("Câu 1. A?") + p("Câu 2. B?"))
    r = import_docx(f, "Toán 7", Session, store)
    assert r.codes == ["C0000001", "C0000002"]
    with pytest.raises(AlreadyImported):
        import_docx(f, "Toán 7", Session, store)
    # cùng nội dung ở chủ đề khác vẫn nhập được và mã tiếp nối, không tái sử dụng
    r2 = import_docx(f, "Toán 8", Session, store)
    assert r2.codes == ["C0000003", "C0000004"]


def test_duplicate_content_flagged_not_dropped(env):
    Session, store, tmp = env
    f = write(tmp, "d.docx", p("Câu 1. Giống nhau?") + p("Câu 2. Giống nhau?"))
    r = import_docx(f, "Toán", Session, store)
    assert r.questions == 2 and r.duplicates == 1
    with Session() as s:
        qs = s.scalars(select(Question).order_by(Question.id)).all()
        assert qs[0].duplicate_of is None and qs[1].duplicate_of == "C0000001"


def test_cli_missing_file_is_clean_error(tmp_path, capsys):
    from eduq.cli import main

    rc = main(["--data", str(tmp_path / "d"), "import", str(tmp_path / "khong-co.docx"), "--topic", "T"])
    assert rc == 1 and "Lỗi" in capsys.readouterr().err


SAMPLE = os.environ.get("EDUQ_SAMPLE_DOCX")


@pytest.mark.skipif(not SAMPLE or not Path(SAMPLE).exists(), reason="đặt EDUQ_SAMPLE_DOCX để chạy với file thật")
def test_real_sample_bo_10_de(env):
    """File thật: 10 đề × (4 trắc nghiệm + 4 tự luận), 397 công thức MathType, 10 hình."""
    Session, store, _ = env
    r = import_docx(SAMPLE, "Toán 7 - GK1", Session, store)
    assert (r.questions, r.multiple_choice, r.free_response) == (80, 40, 40)
    assert r.formulas_total == 397 and r.images == 10
    assert r.formulas_needs_review == 0 and r.questions_needs_review == 6  # 6 câu có hình được chuyển câu
    with Session() as s:
        for q in s.scalars(select(Question)):
            assert q.qtype != "multiple_choice" or [o["label"] for o in q.content["options"]] == list("ABCD")
        text = str([q.content for q in s.scalars(select(Question))])
    assert not re.search(r"(?<!\^)(?<!\^\{)\\\\circ", text)  # không còn dấu độ bị xuất sai
