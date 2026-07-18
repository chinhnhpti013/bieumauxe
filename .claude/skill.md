# Skill: Lập Bộ Hồ Sơ Giám Định Xe Cơ Giới PTI

## Tên skill

`giam-dinh-xe`

## Mô tả

Tự động trích xuất thông tin từ ảnh/PDF tài liệu (màn hình hệ thống PTI, đăng ký xe, GPLX, báo giá, phiếu XMP) rồi điền vào các mẫu biên bản giám định xe cơ giới `.docx`, xuất file hoàn chỉnh vào thư mục `output/`.

---

## Bước 1 — Đọc dữ liệu đầu vào

Đầu vào **chỉ nhận ảnh và PDF** (`.jpg`, `.png`, `.pdf`). Không còn nhập liệu bằng file Excel.

### 1a. Đọc dữ liệu đã nhập (`input/data.json`)

Mọi script đọc dữ liệu qua `pti_common.load_data()` — không tự parse:

```python
from pti_common import load_data

info, phu_tung = load_data()
# info     : dict, mọi giá trị là str, thiếu dữ liệu là "" (không phải None)
# phu_tung : list[{'ten': str, 'phuong_an': str, 'sl': int}]
```

`load_data()` đã tự làm sẵn: ánh xạ mã GĐV → tên + `{SĐT}`, và thêm alias
`Dien_bien_tai_nan`, `ten_chu_xe`, `ten_lai_xe`, `ten_gara_sua_chua`.

File `input/data.json` do `POST /api/save-data` ghi ra từ form web:

```json
{
  "info":     {"bien_so_xe": "14H-042.80", "tien_tt": "125,500,000", ...},
  "phu_tung": [{"ten": "Cản trước", "phuong_an": "Thay thế có thu hồi", "sl": 1}]
}
```

### 1b. Trích xuất từ ảnh / PDF (qua Web App)

Người dùng upload lên giao diện web → bấm **Quét thông tin** → `POST /api/scan-images` gọi Gemini Vision.

**Xử lý PDF trong scan** (`server.py`):
1. Có text layer → `pdfplumber` extract text → gửi dạng text block
2. Scan ảnh (không có text) → `pymupdf` render PNG → gửi dạng ảnh

**Mapping nhãn tài liệu → JSON field** (xem CLAUDE.md mục "SCAN_PROMPT"):

| Tài liệu | Nhãn | Field |
|----------|------|-------|
| GPLX | `Số/No:` | `giay_phep_lai_xe` |
| GPLX | `Họ tên/Full name:` | `lai_xe` |
| GPLX | `Nơi cư trú/Address:` | `dia_chi_lai_xe` |
| GPLX | `Hạng/Class:` | `hang_gplx` |
| GPLX | `Hiệu lực từ ngày/Date:` | `gplx_tu_ngay` |
| GPLX | `Có giá trị đến/Expires:` | `gplx_den_ngay` |
| Phiếu XMP | Số GCN BH | `so_gcn_bh` |
| Phiếu XMP | Phí BH | `phi_bh` |
| Phiếu XMP | Điều kiện bổ sung | `dk_bs` |

**Fallback**: nếu `lai_xe` trống sau scan → server tự copy từ `chu_xe`.

---

## Bước 2 — Xác nhận dữ liệu

Hiển thị bảng tóm tắt các trường đã đọc được. Hỏi người dùng:
- Có trường nào sai/thiếu không?
- Danh sách phụ tùng đúng chưa?

Chờ xác nhận trước khi sang Bước 3.

---

## Bước 3 — Tra mã GĐV → Họ tên + SĐT

`load_data()` đã làm sẵn, dựa trên `GDV_MAP` trong `pti_common.py`: thay `{ma_giam_dinh_vien}` bằng **tên GĐV** và điền `{SĐT}` tương ứng. Không cần làm gì thêm.

> Lưu ý: placeholder trong template là `{ma_giam_dinh_vien}` (không phải `{ten_giam_dinh_vien}`), và `{SĐT}` (chữ hoa, có dấu). Chỉ mẫu Biên bản giám định dùng `{SĐT}`.

Sửa danh sách GĐV thì sửa **cả hai chỗ**: `GDV_MAP` (`pti_common.py`) và `GDV_DEFAULT` (`server.py`, dùng cho dropdown trên web).

---

## Bước 4 — Chuyển số tiền sang chữ

```python
from pti_common import so_thanh_chu

info['tien_tt_chu'] = so_thanh_chu(info.get('tien_tt', '0'))
# "125,500,000" → "Một trăm hai mươi lăm triệu năm trăm nghìn đồng"
```

Hàm nhận cả chuỗi có dấu phân cách (`,` hoặc `.`), không phải số thì trả lại nguyên chuỗi. Đọc đúng chuẩn tiếng Việt: 15 → "mười lăm", 21 → "hai mươi mốt", 25 → "hai mươi lăm". **Đừng tự viết lại hàm này** — bản cũ trong `tao_3_bien_ban.py` từng đọc sai thành "Mười năm", "Hai mươi một".

---

## Bước 5 — Sinh dòng bảng phụ tùng

Phân loại phụ tùng thành 3 nhóm:

```python
from pti_common import danh_sach_phu_tung, phu_tung_thu_hoi

# Điền {ten_phu_tung_1}..{ten_phu_tung_14}, thiếu thì để trống
danh_sach_phu_tung(info, phu_tung, n=14)

# Riêng mẫu Vật tư thu hồi: chỉ phụ tùng "Thay thế có thu hồi"
danh_sach_phu_tung(info, phu_tung_thu_hoi(phu_tung), prefix='ten_pt_thu_hoi', n=13)
```

`phu_tung_thu_hoi()` lọc theo `'thu hồi' in phuong_an` **và** `'không' not in phuong_an` — nên "Thay thế không thu hồi" bị loại đúng.

---

## Bước 6 — Sinh file docx

```python
from pti_common import render, bao_cao_placeholder_con_sot

dst = render('Bảo-lãnh.docx', 'Bao-lanh.docx', info)   # tên file trong assets/ và output/
bao_cao_placeholder_con_sot(dst)                        # in placeholder chưa thay được
```

`render()` lo trọn gói: unpack → `merge_split_placeholders()` → `apply_replacements()`
→ validate XML → pack lại. Chỉ đụng `word/document.xml` + `header*/footer*.xml`,
**không** đụng `styles.xml` / `settings.xml`.

### ⚠️ Vấn đề split placeholder (đã xử lý sẵn)

Word tách `{placeholder}` thành nhiều `<w:r>` run, ví dụ `{ten_phu_tung_2}` →
`{ten_phu_tung_` + `2` + `}`. `render()` đã gọi `merge_split_placeholders()` nên
script không cần tự lo. **Đừng copy hàm này ra script mới** — trước đây nó bị nhân
bản thành 3 biến thể lệch nhau giữa 7 script.

Nếu buộc phải sửa hàm trong `pti_common.py`, nhớ:
- Regex phải là `<w:t(?=[>\s])[^>]*>` — thiếu lookahead sẽ match cả `<w:tab>` → vỡ XML → Word không mở được file.
- Không escape lại `&<>`: text lấy từ `<w:t>` đã được escape sẵn, escape lần nữa sẽ thành `&amp;amp;`.

> **Alias**: `load_data()` đã tự thêm `Dien_bien_tai_nan` (template `Biên-bản-giám-định.docx`
> dùng chữ D hoa), `ten_chu_xe`, `ten_lai_xe`, `ten_gara_sua_chua`.

---

## Bước 7 — Mapping: file template → placeholder dùng

| File template | Placeholder chính |
|--------------|------------------|
| `Biên-bản-giám-định.docx` | Tất cả nhóm A–H + bảng phụ tùng đầy đủ |
| `Vật-tư-thu-hồi.docx` | Nhóm A, C, E + danh sách `thay_co_thu_hoi` |
| `Báo-cáo-sơ-bộ-giám-định.docx` | Nhóm A–F + tóm tắt phụ tùng |
| `Bảo-lãnh.docx` | Nhóm C, D, E, G |
| `TBTN-YCBT.docx` | Nhóm C, D, E, F, G |
| `Xác-nhận-bồi-thường.docx` | Nhóm C, D, E, G + `tien_tt_chu` |
| `CV-Ngân hàng.docx` | Nhóm C, D, E, G, H |
| `Biên-bản-nghiệm-thu.docx` | Nhóm A, C, E + bảng phụ tùng hoàn chỉnh |
| `Cham-dut-khoi-phuc.docx` | Nhóm A, C, D (một phần) + **nhập tay** (xem Bước 9) |

---

## Bước 7b — Mẫu Chấm dứt - Khôi phục (nhập tay ưu tiên)

**Khi nào dùng**: Người dùng gọi "mẫu chấm dứt khôi phục" hoặc cung cấp ảnh thông tin hợp đồng.

**Script chuyên dụng**: `tao_cham_dut_khoi_phuc.py` — sửa khối `NHAP_TAY` ở đầu file trước khi chạy

### Logic xác định ngày khôi phục (từ Phiếu Xác minh phí)

> **Nguồn dữ liệu**: Mục III "Ngày nộp phí trên PMNV" trong Phiếu Xác minh phí (XMP)

**Quy tắc**:
- Tìm dòng đầu tiên có cột **Nội dung = "Doanh thu Thực Thu"**
- **Ngày chứng từ** của dòng đó chính là **ngày khôi phục + 1**
- Suy ra **ngày khôi phục** = ngày chứng từ thực thu lần 1 − 1 ngày

**Ví dụ** (xe 14H-042.80):
| Nội dung | Ngày chứng từ | Ngày hệ thống | Số tiền |
|----------|--------------|--------------|---------|
| Doanh thu bán hàng | 31/12/2025 | 31/12/2025 | 18,551,250 |
| **Doanh thu Thực Thu** | **11/02/2026** | 12/02/2026 | 18,551,250 |

→ Ngày chứng từ thực thu lần 1 = **11/02/2026**
→ Ngày khôi phục = **11/02/2026** (ngày chứng từ thực thu)
→ Ngày khôi phục **+1** = **12/02/2026** (= ngày hệ thống, tức ngày thực sự khôi phục hiệu lực)

Trong script: `ngay_chung_tu_doanh_thu_thuc_thu` điền **ngày chứng từ** (11/02/2026); `ke_tiep_ngay_nop_phi` là hạn kỳ tiếp theo.

### Ánh xạ tự động từ GCN BH / Phiếu XMP

| Placeholder trong docx | Nguồn |
|------------------------|-------|
| `{bien_so_xe}` | GCN BH / `data.json` |
| `{chu_xe}` | GCN BH / `data.json` |
| `{so_gcn_bh}` | GCN BH / `data.json` |
| `{phi_bao_hiem_da_VAT}` | `phi_bh` (đã VAT) |
| `{gcn_bh_gio_phut}` | Giờ phút BH VCX (GCN BH) |
| `{nam_hien_tai}` | Năm hiện tại (tự động) |

### Trường nhập tay / trích xuất từ Phiếu XMP

| Key trong script | Ý nghĩa | Nguồn |
|-----------------|---------|-------|
| `gcn_bh_ngay_cap` | Ngày cấp GCN BH | Phiếu XMP mục I |
| `ngay_thoi_han_bao_hiem_vcx_oto` | Ngày hết hạn BH VCX | GCN BH |
| `gio_phut_thoi_han_bao_hiem_vcx_oto` | Giờ hết hạn BH VCX | GCN BH |
| `ngay/thang/nam_chung_tu_doanh_thu_thuc_thu` | **Ngày chứng từ Doanh thu Thực thu lần 1** | Phiếu XMP mục III |
| `tien_han_thanh_toan` | Số tiền kỳ thanh toán | Phiếu XMP mục III |
| `ngay/thang/nam_han_thanh_toan` | **Hạn thanh toán kỳ 1** (chậm nhất) | Phiếu XMP mục I |
| `ke_tiep_ngay_nop_phi` | Hạn nộp phí kỳ tiếp theo | Phiếu XMP mục I (kỳ 2) |

### Lưu ý kỹ thuật đặc biệt

- Placeholder `{gcn_bh_ ngay_cap}` có **khoảng trắng** bên trong — sau `merge_split_placeholders` vẫn giữ dạng đó. Script xử lý bằng alias `info['gcn_bh_ ngay_cap']`.
- Placeholder `{phi_bao_hiem_da_VAT}` bị split rất nặng (6+ run) — `merge_split_placeholders` sẽ gộp lại đúng.
- Regex kiểm tra placeholder còn sót dùng `r'\{[a-zA-ZÀ-ỹ _][^{}]{1,60}\}'` (có khoảng trắng trong character class) để phát hiện cả `{gcn_bh_ ngay_cap}` nếu chưa thay.
- Tên placeholder trong mẫu docx dùng tên **ngắn** (không có `_lan_1`, `_dong_1`) — cần alias đúng.

### Quy trình khi người dùng cung cấp Phiếu XMP (PDF / ảnh)

1. Đọc Phiếu XMP → tìm dòng "Doanh thu Thực Thu" đầu tiên trong mục III
2. Lấy **Ngày chứng từ** → đó là ngày khôi phục (ngày hệ thống = ngày khôi phục + 1)
3. Lấy hạn thanh toán kỳ 1 từ mục I
4. Xác nhận với người dùng rồi ghi vào script / chạy xuất file

---

## Bước 8 — Chạy và xuất file

Mỗi mẫu là một script `tao_*.py` chạy độc lập (web gọi qua `POST /api/generate`):

| Script | File xuất |
|--------|-----------|
| `tao_bien_ban_gd.py` | `Bien-ban-giam-dinh.docx` |
| `tao_3_bien_ban.py` | `Vat-tu-thu-hoi.docx`, `Xac-nhan-boi-thuong.docx`, `Bien-ban-nghiem-thu.docx` |
| `tao_bao_cao_so_bo.py` | `Bao-cao-so-bo.docx` |
| `tao_bao_lanh.py` | `Bao-lanh.docx` |
| `tao_tbtn_ycbt.py` | `TBTN-YCBT.docx` |
| `tao_cv_ngan_hang.py` | `CV-Ngan-hang.docx` |
| `tao_cham_dut_khoi_phuc.py` | `Cham-dut-khoi-phuc.docx` |

```powershell
python tao_bien_ban_gd.py    # chạy lẻ một mẫu để kiểm tra
```

Cả bộ có chung cấu trúc: `load_data()` → chuẩn bị `info` → `render()`. Xem
`tao_bao_lanh.py` (mẫu đơn giản nhất) làm khuôn khi thêm mẫu mới.

---

## Lưu ý quan trọng

1. **Dùng `python`** (không phải `python3`) trên máy Windows này
2. **Shell là PowerShell** — KHÔNG dùng heredoc `<< 'EOF'`; viết script ra file `.py` rồi chạy `python script.py`
3. **Không dùng `python-docx`** — thay placeholder trực tiếp trong XML để giữ nguyên định dạng gốc
4. **Dùng lại `pti_common.py`**, đừng copy `merge_split_placeholders` / `so_thanh_chu` / `render` vào script mới — các bản nhân bản cũ đã lệch nhau và có bug
5. `render()` đã gọi `merge_split_placeholders()` rồi validate XML bằng `ET.fromstring()` — `ParseError` sẽ dừng ngay với thông báo rõ
6. Placeholder không có dữ liệu → `load_data()` để trống `""` (không giữ nguyên `{...}`)
7. File output lưu vào `output/` với tên ASCII (không dấu) để tránh lỗi encoding trên Windows
8. Encoding mọi file: **UTF-8**
