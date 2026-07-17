# Dự án: Mẫu Biểu Giám Định PTI — Xe Cơ Giới

## Mục tiêu

Tự động lập bộ hồ sơ giám định xe cơ giới PTI từ dữ liệu đầu vào (ảnh màn hình hệ thống / file Excel), điền vào 8 mẫu biên bản `.docx` và xuất file hoàn chỉnh.

## Cấu trúc thư mục

```
mau-bieu-giam-dinh-pti/
├── CLAUDE.md
├── assets/                  # Mẫu biên bản gốc (.docx)
│   ├── Biên-bản-giám-định.docx
│   ├── Vật-tư-thu-hồi.docx
│   ├── Báo-cáo-sơ-bộ-giám-định.docx
│   ├── Bảo-lãnh.docx
│   ├── TBTN-YCBT.docx
│   ├── Xác-nhận-bồi-thường.docx
│   ├── CV-Ngân hàng.docx
│   ├── Biên-bản-nghiệm-thu.docx
│   └── logo-ptisos.png
├── docs/
│   ├── bien-ban-gd-xcg-pti.docx   # Tài liệu mô tả skill/quy trình
│   └── thong_tin_giam_dinh_xe.xlsx # Template nhập liệu
├── input/                   # Dữ liệu đầu vào (ảnh, PDF, Excel)
│   └── *.jpg / *.pdf / *.xlsx
├── output/                  # File xuất ra (tạo tự động)
└── .claude/
    ├── skill.md
    └── settings.json
```

## Dữ liệu đầu vào (thư mục `input/`)

Có thể là:
- **Ảnh màn hình** hệ thống PTI (`.jpg`, `.png`) — AI trích xuất văn bản bằng OCR/vision
- **File Excel** tổng hợp thông tin (`.xlsx`) — đọc trực tiếp qua `openpyxl`
- **File PDF** (`.pdf`) — trích xuất text bằng `pdfplumber` hoặc `pymupdf`

## Template nhập liệu: `docs/thong_tin_giam_dinh_xe.xlsx`

### Sheet "Thông tin" — 8 nhóm placeholder

| Nhóm | Placeholder |
|------|-------------|
| A. Xe | `{bien_so_xe}`, `{hang_xe}`, `{dong_xe}`, `{phien_ban}`, `{nam_sx}`, `{so_loai}`, `{so_khung}`, `{so_may}`, `{trong_tai}`, `{so_cho_ngoi}` |
| B. Giấy tờ | `{giay_phep_luu_hanh}`, `{gplh_tu_ngay}`, `{gplh_den_ngay}`, `{giay_phep_lai_xe}`, `{hang_gplx}`, `{gplx_tu_ngay}`, `{gplx_den_ngay}` |
| C. Chủ/Lái xe | `{chu_xe}`, `{dien_thoai_chu_xe}`, `{dia_chi_chu_xe}`, `{lai_xe}`, `{dien_thoai_lai_xe}`, `{dia_chi_lai_xe}` |
| D. Hợp đồng BH | `{so_hop_dong}`, `{so_gcn_bh}`, `{gcn_bh_tu_ngay}`, `{gcn_bh_den_ngay}`, `{gia_tri_xe}`, `{phi_bh}`, `{dk_bs}` |
| E. Hồ sơ | `{so_ho_so}`, `{ma_giam_dinh_vien}`, `{don_vi_quan_ly}`, `{ngay_giam_dinh}`, `{ngay_vao_gara}` |
| F. Tai nạn | `{ngay_tai_nan}`, `{khoang_gio_tai_nan}`, `{dia_diem_tai_nan}`, `{dien_bien_tai_nan}`, `{nguyen_nhan_tai_nan}` |
| G. Tài chính | `{ngay_hoa_don}`, `{thang_hoa_don}`, `{nam_hoa_don}`, `{tong_thanh_toan_hoa_don}`, `{tien_tt}` |
| H. Gara | `{ten_gara}`, `{so_tk}`, `{ten_ngan_hang}`, `{dia_chi_ngan_hang}` |

### Sheet "Phụ tùng"

| Cột | Mô tả |
|-----|-------|
| Tên phụ tùng | Tên chi tiết cần sửa/thay |
| Phương án | `Thay thế có thu hồi` / `Thay thế không thu hồi` / `Sửa chữa` |
| Số lượng | Số nguyên dương |

## 9 file output (thư mục `output/`)

| Ký hiệu | Tên file xuất | Script |
|---------|--------------|--------|
| A | `Bien-ban-giam-dinh.docx` | `tao_bien_ban_gd.py` |
| B | `Vat-tu-thu-hoi.docx` | `tao_3_bien_ban.py` |
| C | `Bao-cao-so-bo.docx` | — |
| D | `Bao-lanh.docx` | — |
| E | `TBTN-YCBT.docx` | — |
| F | `Xac-nhan-boi-thuong.docx` | `tao_3_bien_ban.py` |
| G | `CV-Ngan-hang.docx` | — |
| H | `Bien-ban-nghiem-thu.docx` | `tao_3_bien_ban.py` |
| I | `Cham-dut-khoi-phuc.docx` | `tao_cham_dut_khoi_phuc.py` |

## Mẫu Chấm dứt - Khôi phục (`Cham-dut-khoi-phuc.docx`)

**Kích hoạt**: Người dùng gọi "mẫu chấm dứt khôi phục" (kèm hoặc không kèm Phiếu Xác minh phí).

### Logic xác định ngày khôi phục (QUAN TRỌNG)

**Nguồn dữ liệu**: Mục III "Ngày nộp phí trên PMNV" trong Phiếu Xác minh phí (XMP).

**Quy tắc**:
- Tìm dòng đầu tiên có cột **Nội dung = "Doanh thu Thực Thu"**
- **Ngày chứng từ** của dòng đó = ngày khôi phục + 1 (tức ngày hệ thống ghi nhận)
- Điền vào placeholder `{ngay/thang/nam_chung_tu_doanh_thu_thuc_thu}` = ngày chứng từ đó

**Ví dụ** (xe 14H-042.80, Phiếu XMP mục III):
- Dòng "Doanh thu Thực Thu" đầu tiên: Ngày chứng từ = **11/02/2026**, Ngày hệ thống = 12/02/2026
- → `ngay_chung_tu_doanh_thu_thuc_thu` = **11**, `thang` = **02**, `nam` = **2026**
- → Ngày khôi phục thực tế = 11/02/2026; ngày hệ thống (khôi phục +1) = 12/02/2026

### Placeholder thực tế trong mẫu docx

| Placeholder | Ý nghĩa | Nguồn |
|-------------|---------|-------|
| `{gcn_bh_ ngay_cap}` | Ngày cấp GCN BH | **Có khoảng trắng** trong tên — Phiếu XMP mục I |
| `{gcn_bh_gio_phut}` | Giờ:phút BH VCX bắt đầu | GCN BH |
| `{ngay_thoi_han_bao_hiem_vcx_oto}` | Ngày hết hạn BH VCX | GCN BH |
| `{gio_phut_thoi_han_bao_hiem_vcx_oto}` | Giờ:phút hết hạn BH VCX | GCN BH |
| `{phi_bao_hiem_da_VAT}` | Phí BH đã VAT | Alias từ `phi_bh` |
| `{nam_hien_tai}` | Năm hiện tại | Tự động |
| `{ngay_chung_tu_doanh_thu_thuc_thu}` | Ngày chứng từ Thực thu lần 1 | Phiếu XMP mục III |
| `{thang_chung_tu_doanh_thu_thuc_thu}` | Tháng chứng từ Thực thu lần 1 | Phiếu XMP mục III |
| `{nam_chung_tu_doanh_thu_thuc_thu}` | Năm chứng từ Thực thu lần 1 | Phiếu XMP mục III |
| `{tien_han_thanh_toan}` | Số tiền kỳ thanh toán | Phiếu XMP mục III |
| `{ngay_han_thanh_toan}` | Ngày hạn TT kỳ 1 | Phiếu XMP mục I |
| `{thang_han_thanh_toan}` | Tháng hạn TT kỳ 1 | Phiếu XMP mục I |
| `{nam_han_thanh_toan}` | Năm hạn TT kỳ 1 | Phiếu XMP mục I |
| `{ngay_thang_nam_han_thanh_toan}` | Chuỗi đầy đủ hạn TT | Tổng hợp tự động |
| `{ke_tiep_ngay_nop_phi}` | Hạn nộp phí kỳ tiếp theo | Phiếu XMP mục I (kỳ 2) |

### Quy trình khi gọi kèm Phiếu XMP

1. Đọc Phiếu XMP (PDF/ảnh) → tìm mục III, dòng "Doanh thu Thực Thu" đầu tiên
2. Lấy **Ngày chứng từ** → điền `ngay/thang/nam_chung_tu_doanh_thu_thuc_thu`
3. Lấy hạn thanh toán kỳ 1 và kỳ 2 từ mục I
4. Xác nhận với người dùng → chạy script xuất file

## Bảng mã Giám định viên

| Mã GĐV | Họ tên |
|--------|--------|
| CHINH05 | Nguyễn Hồng Chinh |
| TUYENLM | Lương Minh Tuyến |
| DUYNT | Nguyễn Thế Duy |
| SONTT | Trần Thanh Sơn |
| VIETNT05 | Nguyễn Tiến Việt |
| HUONGNV | Nguyễn Văn Hướng |
| TUNGHX | Hoàng Xuân Tùng |

## Quy tắc kỹ thuật

- **Python**: dùng `python` (trên máy này, không phải `python3`)
- **Thư viện docx**: xử lý bằng cách unpack XML (zip) → thay thế placeholder → pack lại; KHÔNG dùng `python-docx` để tránh mất định dạng
- **Số tiền sang chữ**: chuyển `{tien_tt}` sang chữ Việt Nam (đồng, trăm, ngàn, triệu, tỷ…)
- **Ngày tháng**: định dạng `DD/MM/YYYY`
- **Encoding**: UTF-8 toàn bộ
- **Placeholder an toàn**: nếu không có dữ liệu → để trống `""` (không giữ nguyên `{placeholder}`)
- **Shell**: PowerShell (Windows) — KHÔNG dùng heredoc `<< 'EOF'`; viết script ra file `.py` rồi chạy `python script.py`

### Excel là nút thắt dữ liệu (QUAN TRỌNG)

Luồng web: form (`index.html`) → `/api/save-excel` → `input/thong_tin_giam_dinh_xe.xlsx` → script sinh docx đọc lại từ Excel.

`save_excel()` chỉ ghi các key **đã có dòng `{placeholder}` sẵn ở cột A** của sheet "Thông tin". Field nào có trong `FIELDS` của `index.html` nhưng thiếu dòng trong Excel sẽ **bị rớt âm thầm, không báo lỗi** → biểu mẫu ra trống.

→ Khi thêm field mới vào form, **bắt buộc thêm dòng tương ứng vào `docs/thong_tin_giam_dinh_xe.xlsx`**.

**Sửa file Excel bằng openpyxl**: các dòng tiêu đề nhóm là merged cell (`A48:C48`…). `insert_rows()` **không** dịch chuyển merged range → ô B/C dòng bên dưới biến thành `MergedCell` và mất dữ liệu. Phải: unmerge tất cả → `insert_rows()` → merge lại với toạ độ đã dịch.

## Xử lý placeholder bị split trong XML (QUAN TRỌNG)

Word đôi khi tách một `{placeholder}` thành nhiều `<w:r>` run liên tiếp trong cùng `<w:p>`. Ví dụ:
- `{ten_phu_tung_2}` → `{ten_phu_tung_` + `2` + `}`
- `{gcn_bh_den_ngay}` → `{gcn_bh_` + `den` + `_ngay}`

**Bắt buộc phải chạy hàm `merge_split_placeholders(xml)` trước khi thay thế.**

Logic hàm này:
1. Duyệt từng `<w:p>` trong XML
2. Tìm tất cả `<w:t>` (regex: `<w:t(?=[>\s])[^>]*>(.*?)</w:t>` — **chú ý lookahead `(?=[>\s])` để KHÔNG khớp `<w:tab>`**)
3. Nối text từ các `<w:t>` lại → kiểm tra có `{placeholder}` hoàn chỉnh không
4. Nếu placeholder bị tách qua nhiều run: gộp text vào run đầu tiên, xóa text các run còn lại

Script mẫu đã hoạt động: `tao_bien_ban_gd.py` (hàm `merge_split_placeholders` + `apply_replacements`).

**Lỗi thường gặp**: Nếu dùng `<w:t[^>]*>` thay vì `<w:t(?=[>\s])[^>]*>` sẽ match cả `<w:tab>` → làm vỡ XML → Word báo lỗi khi mở file.

## Quy trình xử lý chính

1. Đọc dữ liệu đầu vào từ `input/` (Excel hoặc ảnh)
2. Xác nhận / hiển thị dữ liệu đã trích xuất để người dùng kiểm tra
3. Unpack từng file `.docx` trong `assets/` thành XML
4. Chạy `merge_split_placeholders(xml)` để gộp các placeholder bị tách
5. Thay thế tất cả placeholder trong XML
6. Với bảng phụ tùng: điền `{ten_phu_tung_1}` → `{ten_phu_tung_N}` từ sheet "Phụ tùng"
7. Pack lại thành `.docx` → lưu vào `output/`
8. Kiểm tra XML hợp lệ bằng `xml.etree.ElementTree` trước khi báo hoàn tất

## Web App (`server.py` + `index.html`)

Dự án có giao diện web Flask chạy trên Render. Người dùng upload ảnh/PDF → bấm "Quét thông tin" → AI điền form.

### API chính

| Route | Mô tả |
|-------|-------|
| `POST /api/upload-images` | Nhận ảnh (JPG/PNG) và PDF, lưu vào `input/` |
| `POST /api/upload-excel` | Nhận file Excel, đọc và trả dữ liệu |
| `GET /api/input-images` | Liệt kê file trong `input/` (ảnh + PDF) |
| `POST /api/scan-images` | Gọi Claude Vision trích xuất thông tin |
| `POST /api/generate` | Tạo file docx output |

### Xử lý PDF trong scan

1. **Có text layer** (Phiếu XMP): dùng `pdfplumber` extract text → gửi dạng text block vào prompt
2. **Scan ảnh** (không có text): fallback dùng `pymupdf` render từng trang → gửi dạng ảnh PNG

### SCAN_PROMPT — Mapping nhãn tài liệu → JSON field

**Giấy phép lái xe (GPLX)**:

| Nhãn trên thẻ | → JSON field |
|--------------|-------------|
| `Số/No:` | `giay_phep_lai_xe` |
| `Họ tên/Full name:` | `lai_xe` (bắt buộc điền dù trùng chủ xe) |
| `Nơi cư trú/Address:` | `dia_chi_lai_xe` |
| `Hạng/Class:` | `hang_gplx` |
| `Hiệu lực từ ngày/Date:` | `gplx_tu_ngay` |
| `Có giá trị đến/Expires:` | `gplx_den_ngay` |

**Phiếu xác minh phí (XMP)**:

| Nhãn | → JSON field |
|------|-------------|
| Số GCN BH | `so_gcn_bh` |
| Số hợp đồng | `so_hop_dong` |
| Thời hạn BH từ/đến | `gcn_bh_tu_ngay` / `gcn_bh_den_ngay` |
| Phí BH (tổng phí) | `phi_bh` |
| Điều kiện bổ sung | `dk_bs` |

**KHÔNG fallback `lai_xe` ← `chu_xe`**: `lai_xe` / `dia_chi_lai_xe` chỉ lấy từ GPLX hoặc màn hình hệ thống PTI; không có nguồn thì để trống để người dùng tự điền. Trước đây server tự copy từ `chu_xe` khi trống, nhưng chủ xe thường là pháp nhân (vd. `CÔNG TY TNHH PHÚC XUYÊN`) nên sinh ra tên lái xe sai — để trống an toàn hơn là điền bừa.

### Thư viện cần thiết (requirements.txt)

```
flask, openpyxl, gunicorn, mammoth, anthropic, pymupdf, pdfplumber
```

## Quy trình OCR / điền Excel

**Nguồn dữ liệu thực tế** (thư mục `input/`):
- `1.1.jpg`, `1.2.jpg` — Giấy chứng nhận đăng ký xe (2 mặt)
- `2.1.jpg`, `2.2.jpg` — Giấy phép lái xe (2 mặt)
- `3.1.jpg`, `3.2.jpg` — Giấy chứng nhận kiểm định (2 mặt)
- `manhinhb1.jpg` — Màn hình hệ thống PTI tab "Tổn thất - Chi trả" (có số hồ sơ, mã GĐV, diễn biến, gara)
- `manhinhb2.jpg` — Màn hình hệ thống PTI tab "Thông tin giám định" (có thông tin xe, chủ xe, lái xe, ngày GĐ)
- `bao-gia.jpg` — Báo giá sửa chữa từ gara (danh sách phụ tùng)
- `xe_co_gioi_mau_xm_phi.pdf` — Phiếu xác minh phí (số GCN BH, thời hạn BH, phí BH, ĐKBS)

**Lưu ý khi điền Excel**:
- Sheet "Phụ tùng": **chỉ điền cột A** (tên phụ tùng), không chỉnh cột B và C
- `{ma_giam_dinh_vien}` trong biểu mẫu: thay bằng **tên GĐV + SĐT** từ sheet "GĐV"
- Template dùng `{Dien_bien_tai_nan}` (chữ D hoa) — cần alias khi xử lý
- `{SĐT}` (SĐT của GĐV) — lấy từ sheet "GĐV" theo mã GĐV
- Tiền TT lấy từ cột **TIỀN TT** trên màn hình hệ thống (không phải tổng báo giá gara)

## Lưu ý

- Đơn vị quản lý mặc định: **Phòng Giám định và Cứu hộ Quảng Ninh**
- Luôn trả lời và ghi chú bằng **tiếng Việt**
- Không cài thêm thư viện ngoài `openpyxl`, `Pillow`, `pdfplumber`/`pymupdf` nếu không cần thiết
- Validate XML output bằng `ET.fromstring()` — nếu lỗi ParseError thì file docx sẽ không mở được trong Word
