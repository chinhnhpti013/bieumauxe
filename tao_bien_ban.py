# -*- coding: utf-8 -*-
import zipfile, shutil, os, re
from pathlib import Path

# ─── DỮ LIỆU ────────────────────────────────────────────────────────────────

GDV_MAP = {
    "CHINH05":  "Nguyễn Hồng Chinh",
    "TUYENLM":  "Lương Minh Tuyến",
    "DUYNT":    "Nguyễn Thế Duy",
    "SONTT":    "Trần Thanh Sơn",
    "VIETNT05": "Nguyễn Tiến Việt",
    "HUONGNV":  "Nguyễn Văn Hướng",
    "TUNGHX":   "Hoàng Xuân Tùng",
}

data = {
    # A. Xe
    "bien_so_xe":    "14A-894.22",
    "hang_xe":       "MITSUBISHI",
    "dong_xe":       "XFORCE",
    "phien_ban":     "P1 GR1WTGGLVVT",
    "nam_sx":        "2024",
    "so_loai":       "XFORCE P1",
    "so_khung":      "MK2XTGR1WSN005052",
    "so_may":        "4A91KCD6387",
    "trong_tai":     "",
    "so_cho_ngoi":   "5",

    # B. Giấy tờ
    "giay_phep_luu_hanh": "14 001003",
    "gplh_tu_ngay":       "08/07/2024",
    "gplh_den_ngay":      "",
    "giay_phep_lai_xe":   "220153006737",
    "hang_gplx":          "B",
    "gplx_tu_ngay":       "25/06/2025",
    "gplx_den_ngay":      "25/06/2035",

    # C. Chủ/Lái xe
    "chu_xe":           "Đỗ Thu Thủy",
    "dien_thoai_chu_xe": "",
    "dia_chi_chu_xe":   "Tổ 5, Khu 3, Thanh Sơn, Uông Bí, Quảng Ninh",
    "lai_xe":           "Trần Xuân Trường",
    "dien_thoai_lai_xe": "",
    "dia_chi_lai_xe":   "Phường Uông Bí, Tỉnh Quảng Ninh",

    # D. Hợp đồng BH
    "so_hop_dong":    "013OTTN+250003559",
    "so_gcn_bh":      "/ACDT+013OTTN+250003559",
    "gcn_bh_tu_ngay": "05/11/2025",
    "gcn_bh_den_ngay": "05/11/2026",
    "gia_tri_xe":     "600.000.000",
    "phi_bh":         "8.220.000",
    "dk_bs":          "BS02, BS05, BS06",

    # E. Hồ sơ
    "so_ho_so":         "0034938/000-013/0000937/BT/013-KDKVBC/XO/2026",
    "ma_giam_dinh_vien": "CHINH05",
    "don_vi_quan_ly":   "Phòng Giám định và Cứu hộ Quảng Ninh",
    "ngay_giam_dinh":   "31/05/2026",
    "ngay_vao_gara":    "01/06/2026",

    # F. Tai nạn
    "ngay_tai_nan":     "31/05/2026",
    "khoang_gio_tai_nan": "8h15",
    "dia_diem_tai_nan": "Thanh Sơn, Uông Bí, Quảng Ninh",
    "dien_bien_tai_nan": (
        "Khoảng 8h15 ngày 31/05/2026, tại Thanh Sơn - Uông Bí - Quảng Ninh, "
        "xe 14A-894.22 do lái xe Trần Xuân Trường điều khiển, trong khi lùi xe, "
        "do thiếu quan sát đã để xe lùi vào cột điện, khiến xe bị móp đuôi sau xe, gây tai nạn."
    ),
    "nguyen_nhan_tai_nan": "Lái xe lùi xe thiếu quan sát",

    # G. Tài chính
    "tien_tt": "12.540.000",

    # H. Gara
    "ten_gara":        "Công Ty TNHH Huế - Quảng Ninh",
    "so_tk":           "3888698688888",
    "ten_ngan_hang":   "MB Bank",
    "dia_chi_ngan_hang": "",
    "chu_tai_khoan":   "Trần Ngọc Trí",
}

# Tra tên GĐV
ma_gdv = data.get("ma_giam_dinh_vien", "")
data["ten_giam_dinh_vien"] = GDV_MAP.get(ma_gdv, ma_gdv)

# ─── CHUYỂN SỐ TIỀN SANG CHỮ ────────────────────────────────────────────────

def so_thanh_chu(n: int) -> str:
    don_vi = ["", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    don_vi_hang = ["", "mười", "trăm", "nghìn", "mười nghìn", "trăm nghìn",
                   "triệu", "mười triệu", "trăm triệu", "tỷ"]

    if n == 0:
        return "Không đồng"

    def doc_3_so(num, is_first=True):
        tram = num // 100
        chuc = (num % 100) // 10
        don = num % 10
        result = []
        if tram > 0:
            result.append(don_vi[tram] + " trăm")
            if chuc == 0 and don > 0:
                result.append("lẻ")
        elif not is_first and (chuc > 0 or don > 0):
            result.append("không trăm")
            if chuc == 0 and don > 0:
                result.append("lẻ")
        if chuc == 1:
            result.append("mười")
            if don > 0:
                result.append(don_vi[don])
        elif chuc > 1:
            result.append(don_vi[chuc] + " mươi")
            if don == 1:
                result.append("mốt")
            elif don == 5:
                result.append("lăm")
            elif don > 0:
                result.append(don_vi[don])
        elif don > 0 and (tram > 0 or not is_first):
            result.append(don_vi[don])
        elif don > 0:
            result.append(don_vi[don])
        return " ".join(result)

    ty = n // 1_000_000_000
    trieu = (n % 1_000_000_000) // 1_000_000
    nghin = (n % 1_000_000) // 1_000
    le = n % 1_000

    parts = []
    if ty > 0:
        parts.append(doc_3_so(ty, is_first=True) + " tỷ")
    if trieu > 0:
        parts.append(doc_3_so(trieu, is_first=(ty == 0)) + " triệu")
    if nghin > 0:
        parts.append(doc_3_so(nghin, is_first=(ty == 0 and trieu == 0)) + " nghìn")
    if le > 0:
        parts.append(doc_3_so(le, is_first=(ty == 0 and trieu == 0 and nghin == 0)))

    chu = " ".join(parts).strip()
    chu = chu[0].upper() + chu[1:] if chu else "Không"
    return chu + " đồng"

tien_raw = data.get("tien_tt", "0").replace(".", "").replace(",", "").strip()
try:
    data["tien_tt_bang_chu"] = so_thanh_chu(int(tien_raw))
except ValueError:
    data["tien_tt_bang_chu"] = ""

# ─── DANH SÁCH PHỤ TÙNG ────────────────────────────────────────────────────

phu_tung = [
    {"ten": "Gò cốp",              "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Gò cản sau",          "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Gò trong cốp",        "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Đèn hậu bên lái",     "phuong_an": "Thay thế không thu hồi",    "so_luong": 1},
    {"ten": "Sơn cốp ngoài",       "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Sơn cản sau chỗ móp", "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Cản cốp ngoài",       "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Sơn trong cốp",       "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Chạy chỉ cốp trong",  "phuong_an": "Sửa chữa",                  "so_luong": 1},
    {"ten": "Sơn ốp đen cản sau",  "phuong_an": "Sửa chữa",                  "so_luong": 1},
]

thay_co_thu_hoi    = [p for p in phu_tung if p["phuong_an"] == "Thay thế có thu hồi"]
thay_khong_thu_hoi = [p for p in phu_tung if p["phuong_an"] == "Thay thế không thu hồi"]
sua_chua           = [p for p in phu_tung if p["phuong_an"] == "Sửa chữa"]

# Tóm tắt phụ tùng dạng text để điền vào trường đơn giản
def ds_phu_tung_text(ds):
    return "; ".join(f"{p['ten']} (SL: {p['so_luong']})" for p in ds) if ds else ""

data["ds_thay_co_thu_hoi"]    = ds_phu_tung_text(thay_co_thu_hoi)
data["ds_thay_khong_thu_hoi"] = ds_phu_tung_text(thay_khong_thu_hoi)
data["ds_sua_chua"]           = ds_phu_tung_text(sua_chua)
data["ds_phu_tung_tat_ca"]    = ds_phu_tung_text(phu_tung)

# ─── HÀM XỬ LÝ DOCX ────────────────────────────────────────────────────────

def merge_runs_in_xml(text: str) -> str:
    """Gộp các run bị tách khiến placeholder bị chia cắt."""
    # Tìm các cặp </w:t><w:t ...> liền kề trong cùng run và gộp lại
    # Cách đơn giản: làm việc ở cấp paragraph
    return text

def fill_docx(template_path: str, output_path: str, data: dict):
    """Unpack docx, thay placeholder trong XML, pack lại."""
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    tmp = str(output_path) + ".tmp.docx"
    shutil.copy(template_path, tmp)

    with zipfile.ZipFile(tmp, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    for xml_name in [n for n in names if n.endswith(".xml") or n.endswith(".rels")]:
        raw = contents[xml_name]
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            continue

        # Gộp placeholder bị tách bởi XML tags
        # Chiến lược: xóa bỏ các thẻ XML giữa { và } tạm thời
        def merge_placeholder(m):
            inner = re.sub(r'<[^>]+>', '', m.group(0))
            return inner

        text = re.sub(r'\{[^{}]{1,60}\}', merge_placeholder, text)

        for key, val in data.items():
            text = text.replace("{" + key + "}", str(val) if val else "")

        # Xóa placeholder còn sót
        text = re.sub(r"\{[a-z_]+\}", "", text)
        contents[xml_name] = text.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in contents.items():
            zout.writestr(name, content)

    os.remove(tmp)

# ─── CHẠY ────────────────────────────────────────────────────────────────────

ASSETS = Path("assets")
OUTPUT = Path("output")
OUTPUT.mkdir(exist_ok=True)

TEMPLATES = [
    (ASSETS / "Biên-bản-giám-định.docx",      OUTPUT / "Bien-ban-giam-dinh.docx"),
    (ASSETS / "Vật-tư-thu-hồi.docx",          OUTPUT / "Vat-tu-thu-hoi.docx"),
    (ASSETS / "Báo-cáo-sơ-bộ-giám-định.docx", OUTPUT / "Bao-cao-so-bo.docx"),
    (ASSETS / "Bảo-lãnh.docx",                OUTPUT / "Bao-lanh.docx"),
    (ASSETS / "TBTN-YCBT.docx",               OUTPUT / "TBTN-YCBT.docx"),
    (ASSETS / "Xác-nhận-bồi-thường.docx",     OUTPUT / "Xac-nhan-boi-thuong.docx"),
    (ASSETS / "CV-Ngân hàng.docx",            OUTPUT / "CV-Ngan-hang.docx"),
    (ASSETS / "Biên-bản-nghiệm-thu.docx",     OUTPUT / "Bien-ban-nghiem-thu.docx"),
]

print(f"Tiền TT bằng chữ: {data['tien_tt_bang_chu']}")
print(f"GĐV: {data['ten_giam_dinh_vien']}")
print()

errors = []
for template, output in TEMPLATES:
    if not template.exists():
        print(f"⚠  Không tìm thấy template: {template}")
        errors.append(str(template))
        continue
    try:
        fill_docx(str(template), str(output), data)
        print(f"✓  {output.name}")
    except Exception as e:
        print(f"✗  {output.name} — Lỗi: {e}")
        errors.append(str(output))

print()
if errors:
    print(f"Hoàn tất với {len(errors)} lỗi.")
else:
    print("Tất cả file đã được tạo thành công vào thư mục output/")
