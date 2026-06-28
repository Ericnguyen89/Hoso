# Hồ sơ — Tạo bìa hồ sơ lưu trữ tự động

Sinh **bìa hồ sơ lưu trữ** (.docx, khổ A4) hàng loạt từ một tệp dữ liệu **Excel (.xlsx)**.
Giữ nguyên font, khung viền và bố cục của tệp mẫu; mỗi dòng dữ liệu là một bìa, mỗi bìa một trang.

## Tính năng

- Đọc Excel, tự bỏ qua dòng tiêu đề, chuẩn hóa ngày về `dd/mm/yyyy` (thiếu số 0 thì tự thêm,
  chỉ tháng/năm → `mm/yyyy`, dạng dính liền `ddmmyyyy` → `dd/mm/yyyy`).
- Điền dữ liệu theo **nội dung** (không phụ thuộc số thứ tự đoạn) nên bền với thay đổi mẫu.
- **Khối dưới** (Phông số, ngày, số trang…) được neo cố định ở **đáy trang** bằng một text box
  canh đáy riêng — tiêu đề dài bao nhiêu cũng không làm xê dịch.
- Giao diện web hiện đại: kéo-thả Excel → xem trước → tải về `.docx`.

## Cấu trúc cột Excel (sheet đầu)

| Cột | Nội dung |
|-----|----------|
| A | Hồ sơ số |
| B | Nội dung / Tiêu đề hồ sơ |
| C | Thời gian (bắt đầu-kết thúc) |
| D | Số trang |
| E | Số TL |
| F | Thời hạn bảo quản |
| G | Dòng phông (đầu bìa) |

## Cài đặt

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Đặt tệp mẫu `.docx` vào thư mục dự án (mặc định dùng `Sample.docx`).
Có thể đổi thư mục mẫu bằng biến môi trường `BIA_TEMPLATE_DIR`.

## Chạy web

```bash
.venv/bin/python app.py
```

Mở: http://127.0.0.1:5019 (lắng nghe trên `0.0.0.0`, dùng được trong LAN).

## Chạy dòng lệnh (CLI)

```bash
.venv/bin/python tao_bia.py --mau Sample.docx --data "du_lieu.xlsx" --out ket_qua.docx
```

## Tệp chính

- `app.py` — máy chủ Flask (giao diện web).
- `tao_bia.py` — lõi xử lý (đọc Excel, dựng docx); dùng được cả CLI lẫn web.
