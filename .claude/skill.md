# Skill: Lập Bộ Hồ Sơ Giám Định Xe Cơ Giới PTI

## Tên skill

`giam-dinh-xe`

## Mô tả

Tự động trích xuất thông tin từ dữ liệu đầu vào (ảnh màn hình hệ thống PTI hoặc file Excel tổng hợp) rồi điền vào 8 mẫu biên bản giám định xe cơ giới `.docx`, xuất file hoàn chỉnh vào thư mục `output/`.

---

## Bước 1 — Đọc dữ liệu đầu vào

### 1a. Nếu đầu vào là file Excel (`input/*.xlsx`)

```python
import openpyxl

wb = openpyxl.load_workbook("input/<file>.xlsx")

# Sheet "Thông tin": cột A = placeholder, cột B = nhãn, cột C = giá trị
ws_info = wb["Thông tin"]
data = {}
for row in ws_info.iter_rows(min_row=2, values_only=True):
    placeholder, label, value = row[0], row[1], row[2]
    if placeholder and str(placeholder).startswith("{"):
        key = placeholder.strip("{}")
        data[key] = str(value).strip() if value else ""

# Sheet "Phụ tùng": cột A = tên, cột B = phương án, cột C = số lượng
ws_pt = wb["Phụ tùng"]
phu_tung = []
for row in ws_pt.iter_rows(min_row=2, values_only=True):
    ten, phuong_an, so_luong = row[0], row[1], row[2]
    if ten:
        phu_tung.append({
            "ten": str(ten).strip(),
            "phuong_an": str(phuong_an).strip() if phuong_an else "Thay thế không thu hồi",
            "so_luong": int(so_luong) if so_luong else 1
        })
```

### 1b. Nếu đầu vào là ảnh (`input/*.jpg` / `input/*.png`)

- Dùng Claude vision để đọc và trích xuất từng trường thông tin
- Map vào dict `data` cùng cấu trúc với 1a
- Yêu cầu người dùng xác nhận trước khi tiến hành bước tiếp theo

---

## Bước 2 — Xác nhận dữ liệu

Hiển thị bảng tóm tắt các trường đã đọc được. Hỏi người dùng:
- Có trường nào sai/thiếu không?
- Danh sách phụ tùng đúng chưa?

Chờ xác nhận trước khi sang Bước 3.

---

## Bước 3 — Tra mã GĐV → Họ tên + SĐT

Đọc từ sheet "GĐV" trong file Excel (cột A = mã, cột B = tên, cột C = SĐT). Thay `{ma_giam_dinh_vien}` bằng **tên GĐV**, đồng thời cung cấp `{SĐT}` tương ứng.

```python
# Đọc từ sheet GĐV trong Excel (nguồn chính xác nhất)
ws_gdv = wb["GĐV"]
gdv_map = {}
for row in ws_gdv.iter_rows(min_row=2, values_only=True):
    if row[0] and row[1]:
        gdv_map[str(row[0])] = {"ten": str(row[1]), "sdt": str(row[2]) if row[2] else ""}

ma = data.get("ma_giam_dinh_vien", "")
if ma in gdv_map:
    data["ma_giam_dinh_vien"] = gdv_map[ma]["ten"]
    data["SĐT"] = gdv_map[ma]["sdt"]
```

> Lưu ý: placeholder trong template là `{ma_giam_dinh_vien}` (không phải `{ten_giam_dinh_vien}`), và `{SĐT}` (chữ hoa, có dấu).

---

## Bước 4 — Chuyển số tiền sang chữ

```python
def so_thanh_chu(n: int) -> str:
    """Chuyển số nguyên sang chữ Việt Nam (đồng)."""
    # Implement theo chuẩn đọc số tiền Việt Nam
    # Ví dụ: 1_500_000 → "Một triệu năm trăm nghìn đồng"
    ...

data["tien_tt_bang_chu"] = so_thanh_chu(int(data.get("tien_tt", "0").replace(".", "").replace(",", "")))
```

---

## Bước 5 — Sinh dòng bảng phụ tùng

Phân loại phụ tùng thành 3 nhóm:

```python
thay_co_thu_hoi     = [p for p in phu_tung if p["phuong_an"] == "Thay thế có thu hồi"]
thay_khong_thu_hoi  = [p for p in phu_tung if p["phuong_an"] == "Thay thế không thu hồi"]
sua_chua            = [p for p in phu_tung if p["phuong_an"] == "Sửa chữa"]
```

Mỗi nhóm được điền vào bảng phụ tùng tương ứng trong các file docx (xem mapping Bước 6).

---

## Bước 6 — Xử lý docx (unpack XML → fix split → replace → pack)

### ⚠️ Vấn đề split placeholder

Word tách `{placeholder}` thành nhiều `<w:r>` run. **Bắt buộc** phải gộp trước khi replace.

```python
import zipfile, os, re
from xml.etree import ElementTree as ET

def merge_split_placeholders(xml: str) -> str:
    """Gộp các run liên tiếp trong <w:p> mà text nối lại thành {placeholder}."""
    def process_para(para: str) -> str:
        # QUAN TRỌNG: dùng (?=[>\s]) để KHÔNG match <w:tab>, <w:tabs>
        t_re = re.compile(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)', re.DOTALL)
        matches = list(t_re.finditer(para))
        if not matches:
            return para

        texts = [m.group(2) for m in matches]
        combined = ''.join(texts)
        if '{' not in combined:
            return para

        # Bản đồ ký tự → index của match
        char_map = []
        for idx, t in enumerate(texts):
            char_map.extend([idx] * len(t))

        texts = list(texts)
        for ph in re.compile(r'\{[^{}]+\}').finditer(combined):
            s, e = ph.start(), ph.end() - 1
            if s >= len(char_map) or e >= len(char_map):
                continue
            si, ei = char_map[s], char_map[e]
            if si == ei:
                continue  # không bị split
            offset_si  = s - sum(len(texts[k]) for k in range(si))
            offset_ei  = e - sum(len(texts[k]) for k in range(ei)) + 1
            texts[si]  = texts[si][:offset_si] + ph.group(0)
            texts[ei]  = texts[ei][offset_ei:]
            for k in range(si+1, ei):
                texts[k] = ''

        # Ghi lại (duyệt ngược để không lệch offset)
        result = para
        for i in range(len(matches)-1, -1, -1):
            m = matches[i]
            open_tag = '<w:t xml:space="preserve">' if (' ' in texts[i] and 'xml:space' not in m.group(1)) else m.group(1)
            result = result[:m.start()] + open_tag + texts[i] + m.group(3) + result[m.end():]
        return result

    out, pos = [], 0
    for pm in re.finditer(r'<w:p[ >]', xml):
        ps = pm.start()
        pe_idx = xml.find('</w:p>', ps)
        if pe_idx == -1:
            break
        pe = pe_idx + len('</w:p>')
        out.append(xml[pos:ps])
        out.append(process_para(xml[ps:pe]))
        pos = pe
    out.append(xml[pos:])
    return ''.join(out)


def apply_replacements(xml: str, data: dict) -> str:
    for key, val in data.items():
        xml = xml.replace('{' + key + '}', val)
    return xml


def fill_docx(template_path: str, output_path: str, data: dict):
    """Unpack docx, fix split placeholder, thay thế, pack lại."""
    XML_FILES = ('word/document.xml', 'word/header1.xml',
                 'word/footer1.xml', 'word/footer2.xml')

    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_bytes = zin.read(item.filename)
                if item.filename in XML_FILES:
                    xml = data_bytes.decode('utf-8')
                    xml = merge_split_placeholders(xml)
                    xml = apply_replacements(xml, data)
                    data_bytes = xml.encode('utf-8')
                zout.writestr(item, data_bytes)

    # Validate XML — phát hiện lỗi sớm
    with zipfile.ZipFile(output_path, 'r') as z:
        ET.fromstring(z.read('word/document.xml'))
```

> **Alias cần thiết**: template `Biên-bản-giám-định.docx` dùng `{Dien_bien_tai_nan}` (D hoa).
> Thêm `data["Dien_bien_tai_nan"] = data.get("dien_bien_tai_nan", "")` trước khi gọi `fill_docx`.

---

## Bước 7 — Mapping: file template → placeholder dùng

| File template | Placeholder chính |
|--------------|------------------|
| `Biên-bản-giám-định.docx` | Tất cả nhóm A–H + bảng phụ tùng đầy đủ |
| `Vật-tư-thu-hồi.docx` | Nhóm A, C, E + danh sách `thay_co_thu_hoi` |
| `Báo-cáo-sơ-bộ-giám-định.docx` | Nhóm A–F + tóm tắt phụ tùng |
| `Bảo-lãnh.docx` | Nhóm C, D, E, G |
| `TBTN-YCBT.docx` | Nhóm C, D, E, F, G |
| `Xác-nhận-bồi-thường.docx` | Nhóm C, D, E, G + `tien_tt_bang_chu` |
| `CV-Ngân hàng.docx` | Nhóm C, D, E, G, H |
| `Biên-bản-nghiệm-thu.docx` | Nhóm A, C, E + bảng phụ tùng hoàn chỉnh |
| `Cham-dut-khoi-phuc.docx` | Nhóm A, C, D (một phần) + **nhập tay** (xem Bước 9) |

---

## Bước 9 — Mẫu Chấm dứt - Khôi phục (nhập tay ưu tiên)

**Khi nào dùng**: Người dùng gọi "mẫu chấm dứt khôi phục" hoặc cung cấp ảnh thông tin hợp đồng.

**Script chuyên dụng**: `tao_cham_dut_khoi_phuc.py` (hoặc script riêng theo biển số, ví dụ `tao_cham_dut_14H04280.py`)

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
| `{bien_so_xe}` | GCN BH / Excel |
| `{chu_xe}` | GCN BH / Excel |
| `{so_gcn_bh}` | GCN BH / Excel |
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

```python
import os
os.makedirs("output", exist_ok=True)

TEMPLATES = [
    ("assets/Biên-bản-giám-định.docx",        "output/Bien-ban-giam-dinh.docx"),
    ("assets/Vật-tư-thu-hồi.docx",             "output/Vat-tu-thu-hoi.docx"),
    ("assets/Báo-cáo-sơ-bộ-giám-định.docx",   "output/Bao-cao-so-bo.docx"),
    ("assets/Bảo-lãnh.docx",                   "output/Bao-lanh.docx"),
    ("assets/TBTN-YCBT.docx",                  "output/TBTN-YCBT.docx"),
    ("assets/Xác-nhận-bồi-thường.docx",        "output/Xac-nhan-boi-thuong.docx"),
    ("assets/CV-Ngân hàng.docx",               "output/CV-Ngan-hang.docx"),
    ("assets/Biên-bản-nghiệm-thu.docx",        "output/Bien-ban-nghiem-thu.docx"),
]

for template, output in TEMPLATES:
    fill_docx(template, output, data)
    print(f"✓ {output}")
```

---

## Lưu ý quan trọng

1. **Dùng `python`** (không phải `python3`) trên máy Windows này
2. **Shell là PowerShell** — KHÔNG dùng heredoc `<< 'EOF'`; viết script ra file `.py` rồi chạy `python script.py`
3. **Không dùng `python-docx`** — thay placeholder trực tiếp trong XML để giữ nguyên định dạng gốc
4. **Bắt buộc gọi `merge_split_placeholders()`** trước `apply_replacements()` — nếu bỏ qua, nhiều placeholder sẽ không được thay và XML có thể bị vỡ
5. **Regex `<w:t` phải có lookahead `(?=[>\s])`** — nếu dùng `<w:t[^>]*>` sẽ match cả `<w:tab>` → vỡ XML → Word không mở được
6. **Validate sau khi tạo**: gọi `ET.fromstring(...)` trên `word/document.xml`; nếu `ParseError` thì file không dùng được
7. Placeholder còn sót sau khi replace → thay bằng chuỗi rỗng `""`
8. File output lưu vào `output/` với tên ASCII (không dấu) để tránh lỗi encoding trên Windows
9. Encoding mọi file: **UTF-8**
10. **Phụ tùng — chỉ điền cột A** của sheet "Phụ tùng"; không ghi đè cột B (Phương án) và C (Số lượng)
