# eduq – Kho bài tập cho giáo viên

Nhập file Word (công thức **MathType**, hình, chữ) → tách từng câu → lưu vào kho với mã `C0000001…` →
xem lại / xuất lại mà không lỗi công thức. Hiện đã có: **nhập + lưu + xem trước + xuất .docx/PDF**. Chưa có: web UI, nhập đáp án, sửa nội dung câu.

## Cài đặt
```bash
pip install -e '.[dev]'
(cd tools && npm install)                 # KaTeX: kiểm tra + xem trước công thức
gem install mathtype_to_mathml pry        # chỉ cần cho công thức MathType không có TeX nhúng
# cần thêm: pandoc, ruby, node
```

## Dùng
```bash
eduq import de-thi.docx --topic "Toán 7 - GK1"   # in báo cáo import
eduq list --status needs_review                    # câu có công thức cần xem lại
eduq show C0000001                                 # nội dung JSON một câu
eduq preview --out preview                         # trang HTML + KaTeX để xem bằng mắt
eduq export --exam "ĐỀ SỐ 1" --out export          # xuất cả đề → export/de-thi.docx + .pdf
eduq export --codes C0000028,C0000030-C0000032,C0000001 --title "Đề tổng hợp" --name de-tong-hop
eduq export --topic "Toán 7 - GK1" --format docx --paper source   # giữ khổ giấy của file nguồn
```
Xuất đề: chọn bằng `--codes` (đúng thứ tự, nhận khoảng `C…-C…`), `--exam` hoặc `--topic`; `--format docx|pdf|both`;
`--paper a4` (mặc định) hoặc `source`; `--pdf-engine auto|chrome|libreoffice`.
Dữ liệu ở `./data` (SQLite + tài sản theo SHA-256); đổi bằng `--data`.

## Cách xử lý công thức (đã kiểm chứng trên file mẫu 397 công thức)
1. MathType 7 nhúng sẵn mã TeX trong `Equation Native` → dùng trực tiếp (373/384 công thức duy nhất).
2. Không có TeX nhúng → MTEF → MathML (gem Ruby) → LaTeX (pandoc) (11/384).
3. Mọi công thức được **kiểm chéo** với ký tự thật trong MTEF và dựng thử bằng KaTeX nghiêm ngặt;
   không khớp → `needs_review`, vẫn giữ OLE + ảnh WMF gốc.
4. Bảng sửa ký hiệu: MathType xuất `°` thành `\circ` → sửa thành `^{\circ}`; `\prime` → `'`.

## Xuất .docx / PDF
- **.docx**: ghép lại OOXML thô của từng câu nên công thức MathType vẫn là đối tượng OLE như file gốc (sửa được bằng MathType),
  hình giữ nguyên. Style/đánh số lấy từ file nguồn; đánh số lại "Câu 1, 2, …" liền mạch kể cả khi ghép câu từ nhiều file/đề.
  Bỏ đầu trang/watermark. Tự kiểm tra cấu trúc (`check_docx`): r:id treo, thiếu part/content type, id trùng, numId.
- **PDF `chrome`** (mặc định nếu có Chrome/Chromium + `npm install` trong `tools/`): HTML + KaTeX, công thức sắc nét, giống trang web.
  **PDF `libreoffice`**: chuyển từ .docx, công thức hiện bằng ảnh preview gốc của MathType.
- Chạy Chrome dưới root sẽ tự thêm `--no-sandbox`; đặt `EDUQ_CHROME=/đường/dẫn` nếu Chrome không nằm trong PATH.

## Quy ước file Word
Câu hỏi = đoạn đánh số tự động (vd. "Câu 1:") hoặc bắt đầu bằng "Câu N." / "Bài N:". Dòng "ĐỀ SỐ n" mở một đề,
"I. …/II. …" là phần. Dòng "A. … B. … C. … D. …" (một hoặc nhiều dòng) được tách thành phương án.
Watermark/đầu trang không được lưu theo bài.

## Giới hạn đã biết
- Chưa kiểm tra bằng Microsoft Word thật (đã kiểm bằng LibreOffice, python-docx và `check_docx`); nếu Word báo file lỗi, gửi lại file để xử lý.
- Dữ liệu nhập bằng bản cũ (trước khi có xuất) chưa lưu file nguồn → xuất .docx dùng style mặc định; nhập lại để dùng style nguồn.
- Chưa xuất được câu đã sửa nội dung (chưa có tính năng sửa); luôn ghép từ OOXML gốc.
- Công thức Word gốc (OMML) mới được gắn cờ `needs_review`, chưa chuyển sang LaTeX.
- Hộp văn bản (text box) và đối tượng OLE khác ngoài MathType được gắn cờ, chưa hiển thị.
- Bộ quét MTEF dùng để kiểm chéo là heuristic: sai chỉ gây cảnh báo (an toàn), không gây kết quả sai im lặng.
- Chạy test với file thật: `EDUQ_SAMPLE_DOCX=/đường/dẫn/file.docx pytest`.
