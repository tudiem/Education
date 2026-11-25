# eduq – Kho bài tập cho giáo viên

Nhập file Word (công thức **MathType**, hình, chữ) → tách từng câu → lưu vào kho với mã `C0000001…` →
xem lại / xuất lại mà không lỗi công thức. Hiện đã có: **nhập + lưu + xem trước**. Chưa có: web UI, xuất .docx/PDF, nhập đáp án.

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
```
Dữ liệu ở `./data` (SQLite + tài sản theo SHA-256); đổi bằng `--data`.

## Cách xử lý công thức (đã kiểm chứng trên file mẫu 397 công thức)
1. MathType 7 nhúng sẵn mã TeX trong `Equation Native` → dùng trực tiếp (373/384 công thức duy nhất).
2. Không có TeX nhúng → MTEF → MathML (gem Ruby) → LaTeX (pandoc) (11/384).
3. Mọi công thức được **kiểm chéo** với ký tự thật trong MTEF và dựng thử bằng KaTeX nghiêm ngặt;
   không khớp → `needs_review`, vẫn giữ OLE + ảnh WMF gốc.
4. Bảng sửa ký hiệu: MathType xuất `°` thành `\circ` → sửa thành `^{\circ}`; `\prime` → `'`.

## Quy ước file Word
Câu hỏi = đoạn đánh số tự động (vd. "Câu 1:") hoặc bắt đầu bằng "Câu N." / "Bài N:". Dòng "ĐỀ SỐ n" mở một đề,
"I. …/II. …" là phần. Dòng "A. … B. … C. … D. …" (một hoặc nhiều dòng) được tách thành phương án.
Watermark/đầu trang không được lưu theo bài.

## Giới hạn đã biết
- Công thức Word gốc (OMML) mới được gắn cờ `needs_review`, chưa chuyển sang LaTeX.
- Hộp văn bản (text box) và đối tượng OLE khác ngoài MathType được gắn cờ, chưa hiển thị.
- Bộ quét MTEF dùng để kiểm chéo là heuristic: sai chỉ gây cảnh báo (an toàn), không gây kết quả sai im lặng.
- Chạy test với file thật: `EDUQ_SAMPLE_DOCX=/đường/dẫn/file.docx pytest`.
