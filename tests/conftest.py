import io
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def make_docx(body_xml: str, numbering_xml: str | None = None) -> bytes:
    """Tạo .docx tối thiểu trong bộ nhớ (không cần MathType) để test parser/importer."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body_xml}</w:body></w:document>')
        if numbering_xml:
            z.writestr("word/numbering.xml", f'<?xml version="1.0"?><w:numbering {W}>{numbering_xml}</w:numbering>')
    return buf.getvalue()


def p(text: str, *, bold=False, num_id: str | None = None) -> str:
    ppr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>' if num_id else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
