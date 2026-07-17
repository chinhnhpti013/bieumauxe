# Dự án: Mẫu Biểu Giám Định PTI — Xe Cơ Giới

## Mục tiêu

Tự động lập bộ hồ sơ giám định xe cơ giới PTI từ ảnh/PDF tài liệu (ảnh màn hình hệ thống, đăng ký xe, GPLX, báo giá, phiếu XMP), điền vào các mẫu biên bản `.docx` và xuất file hoàn chỉnh.

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
│   └── bien-ban-gd-xcg-pti.docx   # Tài liệu mô tả skill/quy trình
├── input/                   # Dữ liệu đầu vào (ảnh, PDF) + data.json
│   ├── *.jpg / *.png / *.pdf
│   └── data.json            # Dữ liệu form (server ghi ra, script đọc vào)
├── output/                  # File xuất ra (tạo tự động)
├── pti_common.py            # Module dùng chung của mọi script tao_*.py
├── server.py                # Flask backend
├── index.html               # Giao diện web
└── .claude/
    ├── skill.md
    └── settings.json
```

## Dữ liệu đầu vào (thư mục `input/`)

**Chỉ nhận ảnh và PDF** — không còn nhập liệu bằng file Excel:
- **Ảnh** (`.jpg`, `.png`) — đăng ký xe, GPLX, màn hình hệ thống PTI, báo giá gara. AI trích xuất bằng vision.
- **File PDF** (`.pdf`) — phiếu XMP, GCN BH. Có text layer thì dùng `pdfplumber`; PDF scan ảnh thì render bằng `pymupdf` rồi đưa qua vision.

## Dữ liệu trung gian: `input/data.json`

Người dùng nhập/sửa trên form web → `POST /api/save-data` ghi ra `input/data.json` → các script `tao_*.py` đọc lại qua `pti_common.load_data()`.

```json
{
  "info":     {"bien_so_xe": "14H-042.80", "tien_tt": "125,500,000", ...},
  "phu_tung": [{"ten": "Cản trước", "phuong_an": "Thay thế có thu hồi", "sl": 1}]
}
```

Thêm field mới **chỉ cần** thêm vào `FIELDS` trong `index.html` — JSON không có schema cố định nên dữ liệu không bị rớt. (Trước đây dùng Excel làm trung gian: thiếu dòng `{placeholder}` trong sheet là field bị rớt âm thầm, đã gây lỗi mất `{tien_tt}`.)

### 8 nhóm placeholder

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

### Phụ tùng (`phu_tung[]`)

| Khoá | Mô tả |
|------|-------|
| `ten` | Tên chi tiết cần sửa/thay |
| `phuong_an` | `Thay thế có thu hồi` / `Thay thế không thu hồi` / `Sửa chữa` |
| `sl` | Số lượng, số nguyên dương |

## 9 file output (thư mục `output/`)

| Ký hiệu | Tên file xuất | Script |
|---------|--------------|--------|
| A | `Bien-ban-giam-dinh.docx` | `tao_bien_ban_gd.py` |
| B | `Vat-tu-thu-hoi.docx` | `tao_3_bien_ban.py` |
| C | `Bao-cao-so-bo.docx` | `tao_bao_cao_so_bo.py` |
| D | `Bao-lanh.docx` | `tao_bao_lanh.py` |
| E | `TBTN-YCBT.docx` | `tao_tbtn_ycbt.py` |
| F | `Xac-nhan-boi-thuong.docx` | `tao_3_bien_ban.py` |
| G | `CV-Ngan-hang.docx` | `tao_cv_ngan_hang.py` |
| H | `Bien-ban-nghiem-thu.docx` | `tao_3_bien_ban.py` |
| I | `Cham-dut-khoi-phuc.docx` | `tao_cham_dut_khoi_phuc.py` |

Khoá `SCRIPT_MAP` trong `server.py`: `A`, `BFH` (một script sinh cả 3), `C`, `D`, `E`, `G`, `I`.

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

## Module dùng chung `pti_common.py` (QUAN TRỌNG)

Mọi script `tao_*.py` đều dùng module này — **không copy lại các hàm bên dưới vào script mới**.

| Hàm | Công dụng |
|-----|-----------|
| `load_data()` | Đọc `input/data.json` → `(info, phu_tung)`. Ép mọi giá trị về `str`, `None` → `""`. Tự ánh xạ mã GĐV → tên + SĐT, và thêm alias (`Dien_bien_tai_nan`, `ten_chu_xe`, `ten_lai_xe`, `ten_gara_sua_chua`). |
| `danh_sach_phu_tung(info, pt, prefix, n)` | Điền `{prefix}_1`…`{prefix}_n`, thiếu thì để trống. |
| `phu_tung_thu_hoi(pt)` | Lọc phụ tùng phương án "Thay thế có thu hồi". |
| `so_thanh_chu(so)` | Số tiền → chữ tiếng Việt (đúng cách đọc: 15 → "mười lăm", 21 → "hai mươi mốt"). |
| `merge_split_placeholders(xml)` | Gộp placeholder bị Word tách qua nhiều `<w:r>`. |
| `apply_replacements(xml, info)` | Thay `{key}` → giá trị. |
| `render(asset, output, info)` | Unpack docx → merge → replace → validate XML → pack. Trả về đường dẫn. |
| `bao_cao_placeholder_con_sot(dst)` | In các placeholder chưa thay được. |

Script điển hình chỉ còn vài dòng:

```python
from pti_common import load_data, render, bao_cao_placeholder_con_sot

info, _ = load_data()
dst = render('Bảo-lãnh.docx', 'Bao-lanh.docx', info)
bao_cao_placeholder_con_sot(dst)
```

`render()` chỉ đụng vào `word/document.xml` + `header*/footer*.xml`, **cố ý không đụng** `styles.xml` / `settings.xml`.

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

Hàm này nằm trong `pti_common.py` và `render()` đã gọi sẵn — script không cần tự gọi.

**Lỗi thường gặp**: Nếu dùng `<w:t[^>]*>` thay vì `<w:t(?=[>\s])[^>]*>` sẽ match cả `<w:tab>` → làm vỡ XML → Word báo lỗi khi mở file.

## Quy trình xử lý chính

1. Upload ảnh/PDF vào `input/` → `POST /api/scan-images` để AI trích xuất
2. Người dùng kiểm tra / sửa trên form → `POST /api/save-data` ghi `input/data.json`
3. `POST /api/generate` chạy các script `tao_*.py` (mỗi script một tiến trình riêng)
4. Script gọi `load_data()` → chuẩn bị `info` → `render()` → `output/*.docx`
5. `render()` tự lo: merge placeholder bị tách → thay thế → validate XML → pack lại

## Web App (`server.py` + `index.html`)

Dự án có giao diện web Flask chạy trên Render. Người dùng upload ảnh/PDF → bấm "Quét thông tin" → AI điền form.

### API chính

| Route | Mô tả |
|-------|-------|
| `POST /api/upload-images` | Nhận ảnh (JPG/PNG) và PDF, lưu vào `input/` |
| `GET /api/input-images` | Liệt kê file trong `input/` (ảnh + PDF) |
| `POST /api/scan-images` | Gọi Gemini Vision trích xuất thông tin |
| `POST /api/save-data` | Ghi dữ liệu form xuống `input/data.json` |
| `POST /api/generate` | Chạy script `tao_*.py` → tạo file docx |
| `GET /api/gdv-list` | Danh sách giám định viên (`GDV_DEFAULT`) |
| `POST /api/reset` | Xoá sạch `input/` và `output/` |

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
flask, python-dotenv, gunicorn, mammoth, google-genai, pymupdf, pdfplumber
```

Các script `tao_*.py` **chỉ dùng thư viện chuẩn** (`json`, `re`, `zipfile`) — không phụ thuộc thư viện ngoài nào.

## Quy trình OCR / nhập liệu

**Nguồn dữ liệu thực tế** (thư mục `input/`):
- `1.1.jpg`, `1.2.jpg` — Giấy chứng nhận đăng ký xe (2 mặt)
- `2.1.jpg`, `2.2.jpg` — Giấy phép lái xe (2 mặt)
- `3.1.jpg`, `3.2.jpg` — Giấy chứng nhận kiểm định (2 mặt)
- `manhinhb1.jpg` — Màn hình hệ thống PTI tab "Tổn thất - Chi trả" (có số hồ sơ, mã GĐV, diễn biến, gara)
- `manhinhb2.jpg` — Màn hình hệ thống PTI tab "Thông tin giám định" (có thông tin xe, chủ xe, lái xe, ngày GĐ)
- `bao-gia.jpg` — Báo giá sửa chữa từ gara (danh sách phụ tùng)
- `xe_co_gioi_mau_xm_phi.pdf` — Phiếu xác minh phí (số GCN BH, thời hạn BH, phí BH, ĐKBS)

**Lưu ý khi nhập liệu**:
- `{ma_giam_dinh_vien}` trong biểu mẫu: `load_data()` tự thay mã bằng **tên GĐV**; `{SĐT}` lấy từ `GDV_MAP` trong `pti_common.py` (chỉ mẫu Biên bản giám định dùng `{SĐT}`)
- Template dùng `{Dien_bien_tai_nan}` (chữ D hoa) — `load_data()` đã tự alias
- Tiền TT lấy từ cột **TIỀN TT** trên màn hình hệ thống (không phải tổng báo giá gara)

## Lưu ý

- Đơn vị quản lý mặc định: **Phòng Giám định và Cứu hộ Quảng Ninh**
- Luôn trả lời và ghi chú bằng **tiếng Việt**
- Không cài thêm thư viện ngoài `pdfplumber`/`pymupdf` nếu không cần thiết
- Validate XML output bằng `ET.fromstring()` — nếu lỗi ParseError thì file docx sẽ không mở được trong Word
