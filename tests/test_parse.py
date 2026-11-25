from conftest import make_docx, p

from eduq.package import DocxError, DocxPackage
from eduq.parse import parse_document

NUMBERING = (
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="Câu %1:"/></w:lvl></w:abstractNum>'
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>'
    '<w:num w:numId="10"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="11"><w:abstractNumId w:val="1"/></w:num>'
)


def parse(body, numbering=None):
    return parse_document(DocxPackage(make_docx(body, numbering)))


def test_text_markers_and_exam_sections():
    doc = parse(
        p("ĐỀ SỐ 1") + p("I. TRẮC NGHIỆM (1 điểm)") +
        p("Câu 1. Số nào là số chẵn?") + p("A. 1          B. 2          C. 3          D. 5") +
        p("II. PHẦN TỰ LUẬN") + p("Câu 2. (3,0 điểm) Tính 2+2.") +
        p("ĐỀ SỐ 2") + p("Câu 1. Hết.")
    )
    qs = doc.questions
    assert [(q.exam, q.number) for q in qs] == [("ĐỀ SỐ 1", 1), ("ĐỀ SỐ 1", 2), ("ĐỀ SỐ 2", 1)]
    assert qs[0].qtype == "multiple_choice" and qs[0].section.startswith("I. TRẮC NGHIỆM")
    assert [o["label"] for o in qs[0].content["options"]] == list("ABCD")
    assert qs[0].content["options"][1]["inlines"][0]["text"] == "2"
    assert qs[1].qtype == "free_response" and qs[1].points == 3.0 and qs[1].section.startswith("II.")


def test_auto_numbered_list_is_question_bullets_are_not():
    doc = parse(
        p("Tính 1+1?", num_id="10") + p("A. 1  B. 2  C. 3  D. 4") +
        p("Gạch đầu dòng", num_id="11") +
        p("Tính 2+2?", num_id="10"),
        NUMBERING,
    )
    assert len(doc.questions) == 2
    assert doc.questions[0].qtype == "multiple_choice"
    # gạch đầu dòng là phần tiếp theo của câu 1 (sau phương án → "after"), không phải câu mới
    c = doc.questions[0].content
    assert any("Gạch đầu dòng" in str(b) for b in c["after"])
    assert "Gạch đầu dòng" not in str(c["options"])


def test_options_must_be_in_order_else_free_response():
    doc = parse(p("Câu 1. Hỏi?") + p("A. x  C. y  B. z  D. w"))
    assert doc.questions[0].qtype == "free_response"


def test_marks_and_raw_ooxml_kept():
    doc = parse(p("Câu 1. Đậm", bold=True) + p("nội dung"))
    q = doc.questions[0]
    assert q.content["stem"][0]["inlines"][0] == {"type": "text", "text": "Đậm", "marks": {"b": True}}
    # nhãn "Câu 1." không nằm trong nội dung lưu, nhưng OOXML thô vẫn giữ nguyên
    assert "Câu 1." in q.raw_ooxml and "<w:b" in q.raw_ooxml


def test_floating_offset_numbers_not_in_text():
    anchor = '<w:p><w:r><w:t>Câu 1. Có hình</w:t></w:r><w:r><wp:posOffset xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">403034518224500</wp:posOffset></w:r></w:p>'
    q = parse(anchor).questions[0]
    assert q.content["stem"][0]["inlines"][0]["text"] == "Có hình"
    assert "403034518224500" not in str(q.content)


def test_invalid_docx_rejected():
    import pytest
    with pytest.raises(DocxError):
        DocxPackage(b"khong phai zip")
