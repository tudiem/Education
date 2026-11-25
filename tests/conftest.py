import io
import zipfile

W = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360000002000001e221bc330000000049454e44ae426082"
)


def make_docx(body_xml: str, numbering_xml: str | None = None, images: dict[str, bytes] | None = None) -> bytes:
    """Tạo .docx tối thiểu trong bộ nhớ (không cần MathType) để test parser/importer."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body_xml}</w:body></w:document>')
        if images:  # {"rId5": png_bytes} → word/media/imageN.png + rels
            rels = "".join(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{i}.png"/>'
                for i, rid in enumerate(images, 1)
            )
            z.writestr("word/_rels/document.xml.rels", f'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
            for i, data in enumerate(images.values(), 1):
                z.writestr(f"word/media/image{i}.png", data)
        if numbering_xml:
            z.writestr("word/numbering.xml", f'<?xml version="1.0"?><w:numbering {W}>{numbering_xml}</w:numbering>')
    return buf.getvalue()


def p(text: str, *, bold=False, num_id: str | None = None) -> str:
    ppr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>' if num_id else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def p_split(*runs: str) -> str:
    """Đoạn gồm nhiều run (để thử nhãn "Câu " + "3." bị tách)."""
    return "<w:p>" + "".join(f'<w:r><w:t xml:space="preserve">{r}</w:t></w:r>' for r in runs) + "</w:p>"


def p_image(text: str, rid: str, *, num_id: str | None = None, floating: bool = False, crop_top: int = 0) -> str:
    ppr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>' if num_id else ""
    kind = "anchor" if floating else "inline"
    src = f'<a:srcRect t="{crop_top}"/>' if crop_top else ""
    drawing = (
        f'<w:r><w:drawing><wp:{kind}><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Hinh"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:blipFill><a:blip r:embed="{rid}"/>{src}</pic:blipFill></pic:pic></a:graphicData></a:graphic>'
        f"</wp:{kind}></w:drawing></w:r>"
    )
    return f'<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r>{drawing}</w:p>'
