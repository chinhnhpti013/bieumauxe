# -*- coding: utf-8 -*-
import zipfile, re, os, shutil

ASSET = r"d:\MCP\Claude Code\mau bieu giam dinh pti\assets\CV-Ngân hàng.docx"
OUTPUT = r"d:\MCP\Claude Code\mau bieu giam dinh pti\output\CV-Ngan-hang.docx"

os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

def so_sang_chu(so_str):
    so = int(so_str.replace(",", "").replace(".", ""))
    don_vi = ["", "nghìn", "triệu", "tỷ"]
    chu_so = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def doc_3(n):
        if n == 0:
            return ""
        tram = n // 100
        chuc = (n % 100) // 10
        dv = n % 10
        r = ""
        if tram > 0:
            r += chu_so[tram] + " trăm"
        if chuc == 0:
            if dv > 0 and tram > 0:
                r += " lẻ " + chu_so[dv]
            elif dv > 0:
                r += chu_so[dv]
        elif chuc == 1:
            r += " mười"
            if dv == 5: r += " lăm"
            elif dv > 0: r += " " + chu_so[dv]
        else:
            r += " " + chu_so[chuc] + " mươi"
            if dv == 1: r += " mốt"
            elif dv == 5: r += " lăm"
            elif dv > 0: r += " " + chu_so[dv]
        return r.strip()

    if so == 0:
        return "Không đồng"
    groups = []
    tmp = so
    while tmp > 0:
        groups.append(tmp % 1000)
        tmp //= 1000
    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if g > 0:
            txt = doc_3(g)
            if don_vi[i]:
                txt += " " + don_vi[i]
            parts.append(txt)
    result = " ".join(parts).strip()
    return result[0].upper() + result[1:] + " đồng"


def merge_split_placeholders(xml):
    def process_para(m):
        para = m.group(0)
        t_pat = re.compile(r'<w:t(?=[>\s])[^>]*>(.*?)</w:t>', re.DOTALL)
        ts = list(t_pat.finditer(para))
        if not ts:
            return para
        full = "".join(x.group(1) for x in ts)
        if "{" not in full or "}" not in full:
            return para
        has_split = any(("{" in x.group(1) and "}" not in x.group(1)) or
                        ("}" in x.group(1) and "{" not in x.group(1)) for x in ts)
        if not has_split:
            return para
        # Gộp vào run đầu
        new_para = para
        first = ts[0]
        safe = full.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # Khôi phục lại & đã có trong XML gốc (chỉ encode text mới)
        # full_text đã là text thuần (đã được decode bởi regex) nên cần encode lại
        new_t = re.sub(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)',
                       lambda mm: mm.group(1) + safe + mm.group(3),
                       first.group(0), count=1, flags=re.DOTALL)
        new_para = new_para.replace(first.group(0), new_t, 1)
        for t in ts[1:]:
            cleared = re.sub(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)',
                             lambda mm: mm.group(1) + mm.group(3),
                             t.group(0), count=1, flags=re.DOTALL)
            new_para = new_para.replace(t.group(0), cleared, 1)
        return new_para

    return re.sub(r'<w:p[ >].*?</w:p>', process_para, xml, flags=re.DOTALL)


def apply_replacements(xml, data):
    for k, v in data.items():
        xml = xml.replace("{" + k + "}", v)
    return xml


TIEN_CHU = so_sang_chu("6432367")

data = {
    "ten_ngan_hang": "Ngân hàng TMCP Việt Nam Thịnh Vượng (VPBank)",
    "dia_chi_ngan_hang": "Số 31-33, Phạm Ngũ Lão, phường Gia Viễn, thành phố Hải Phòng",
    "bien_so_xe": "14C-402.84",
    "chu_xe": "CÔNG TY CỔ PHẦN ĐẦU TƯ XÂY DỰNG UÔNG BÍ",
    "ten_chu_xe": "CÔNG TY CỔ PHẦN ĐẦU TƯ XÂY DỰNG UÔNG BÍ",
    "dia_chi_chu_xe": "Số 513 đường Quang Trung, Phường Uông Bí, Quảng Ninh",
    "so_khung": "RL2UMFC30RCR74868",
    "so_gcn_bh": "25002.1.820",
    "gcn_bh_tu_ngay": "23/07/2025",
    "gcn_bh_den_ngay": "23/07/2026",
    "ngay_tai_nan": "08/06/2026",
    "lai_xe": "Đỗ Quang Thịnh",
    "dia_diem_tai_nan": "Uông Bí, Quảng Ninh",
    "so_tien_thiet_hai_so": "6.432.367",
    "so_tien_thiet_hai_chu": TIEN_CHU,
}

print("Số tiền bằng chữ:", TIEN_CHU)

with zipfile.ZipFile(ASSET, 'r') as zin:
    all_files = zin.namelist()
    file_contents = {f: zin.read(f) for f in all_files}

xml = file_contents['word/document.xml'].decode('utf-8')
xml = merge_split_placeholders(xml)
xml = apply_replacements(xml, data)

import xml.etree.ElementTree as ET
try:
    ET.fromstring(xml)
    print("XML hợp lệ.")
except ET.ParseError as e:
    print("LỖI XML:", e)
    raise

file_contents['word/document.xml'] = xml.encode('utf-8')

with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for fname, content in file_contents.items():
        zout.writestr(fname, content)

print(f"Đã tạo: {OUTPUT}")
