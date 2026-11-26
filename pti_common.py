# -*- coding: utf-8 -*-
"""Tiện ích dùng chung cho các script sinh biểu mẫu.

Dữ liệu đầu vào là `input/data.json`, do `/api/save-data` của server ghi ra:

    {
      "info":     {"bien_so_xe": "14H-042.80", ...},
      "phu_tung": [{"ten": "Cản trước", "phuong_an": "Thay thế có thu hồi", "sl": 1}, ...]
    }
"""
import json
import os
import re
import zipfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(BASE_DIR, 'input', 'data.json')
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# Phần XML có thể chứa placeholder. Hiện mọi mẫu chỉ đặt placeholder trong
# document.xml, nhưng vẫn quét header/footer phòng khi mẫu sau này dùng tới.
# Cố ý KHÔNG đụng styles.xml / settings.xml — sửa ở đó chỉ có hại.
_XML_PART_RE = re.compile(r'^word/(document|header\d*|footer\d*)\.xml$')

GDV_MAP = {
    'CHINH05':  {'ten': 'Nguyễn Hồng Chinh', 'sdt': '0888955673'},
    'TUYENLM':  {'ten': 'Lương Minh Tuyến',  'sdt': ''},
    'DUYNT':    {'ten': 'Nguyễn Thế Duy',    'sdt': ''},
    'SONTT':    {'ten': 'Trần Thanh Sơn',    'sdt': ''},
    'VIETNT05': {'ten': 'Nguyễn Tiến Việt',  'sdt': ''},
    'HUONGNV':  {'ten': 'Nguyễn Văn Hướng',  'sdt': ''},
    'TUNGHX':   {'ten': 'Hoàng Xuân Tùng',   'sdt': ''},
}


# ── Đọc dữ liệu ───────────────────────────────────────────────────────
def load_data():
    """Đọc input/data.json → (info, phu_tung).

    Mọi giá trị đều được ép về str; None → "" để placeholder không có dữ liệu
    được xoá khỏi văn bản thay vì in nguyên `{ten_placeholder}` ra file Word.
    """
    if not os.path.exists(DATA_JSON):
        raise SystemExit(
            f'Không tìm thấy {DATA_JSON}.\n'
            'Hãy nhập thông tin trên web rồi bấm "Lưu dữ liệu" trước khi tạo biểu mẫu.'
        )

    with open(DATA_JSON, encoding='utf-8') as f:
        data = json.load(f)

    info = {k: ('' if v is None else str(v)) for k, v in data.get('info', {}).items()}

    phu_tung = []
    for idx, pt in enumerate(data.get('phu_tung', []), start=1):
        ten = str(pt.get('ten', '')).strip()
        if not ten:
            continue
        phu_tung.append({
            # 'stt' = số thứ tự trên tab Phụ tùng ở Bước 2, giữ nguyên kể cả khi
            # danh sách bị lọc (vd Vật tư thu hồi chỉ lấy hạng mục có thu hồi)
            'stt': idx,
            'ten': ten,
            'phuong_an': str(pt.get('phuong_an', '') or ''),
            'muc_do': str(pt.get('muc_do', '') or ''),
            'sl': pt.get('sl') or 1,
        })

    # Mã GĐV → tên + SĐT (biểu mẫu in tên người, không in mã)
    gdv = GDV_MAP.get(info.get('ma_giam_dinh_vien', ''))
    info['ma_giam_dinh_vien'] = gdv['ten'] if gdv else info.get('ma_giam_dinh_vien', '')
    info['SĐT'] = gdv['sdt'] if gdv else ''

    # Alias: một số mẫu dùng tên placeholder khác cho cùng dữ liệu
    info['Dien_bien_tai_nan'] = info.get('dien_bien_tai_nan', '')
    info['ten_chu_xe'] = info.get('chu_xe', '')
    info['ten_lai_xe'] = info.get('lai_xe', '')
    info['ten_gara_sua_chua'] = info.get('ten_gara', '')
    info['dia_chi_hoa_don_nguoi_ban'] = info.get('dia_chi_gara', '')

    # Tách ngày/tháng/năm cho các mẫu in rời 3 ô
    #   {ngay_bb_gd}    Biên bản giám định      ← Ngày tai nạn
    #   {ngay_hoa_don}  Vật tư thu hồi, Xác nhận bồi thường,
    #                   phần đầu Biên bản nghiệm thu
    #                                            ← Ngày nghiệm thu
    #   {ngay_bao_gia}  dòng "Căn cứ Hợp đồng/Báo giá sửa chữa" của Nghiệm thu
    #                                            ← Ngày báo giá
    for prefix, nguon in (('bb_gd', 'ngay_tai_nan'),
                          ('hoa_don', 'ngay_nghiem_thu'),
                          ('bao_gia', 'ngay_bao_gia')):
        ngay, thang, nam = tach_ngay(info.get(nguon, ''))
        info['ngay_' + prefix] = ngay
        info['thang_' + prefix] = thang
        info['nam_' + prefix] = nam

    return info, phu_tung


def tach_ngay(chuoi):
    """'14/12/2028' → ('14', '12', '2028'). Không đọc được → ('', '', '')."""
    m = re.match(r'^\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\s*$', str(chuoi or ''))
    if not m:
        return '', '', ''
    return '%02d' % int(m.group(1)), '%02d' % int(m.group(2)), m.group(3)


def danh_sach_phu_tung(info, phu_tung, prefix='ten_phu_tung', n=17, stt=False):
    """Điền {prefix}_1..{prefix}_n; thiếu thì để trống.

    stt=True thì điền luôn {stt1}..{sttN} bằng số thứ tự GỐC của hạng mục trên
    tab Phụ tùng ở Bước 2 (không đánh lại từ 1 khi danh sách đã bị lọc).
    """
    for i in range(1, n + 1):
        pt = phu_tung[i - 1] if i - 1 < len(phu_tung) else None
        info[f'{prefix}_{i}'] = pt['ten'] if pt else ''
        if stt:
            info[f'stt{i}'] = str(pt.get('stt', i)) if pt else ''


def bang_bien_ban_gd(info, phu_tung, n=17):
    """Bảng hạng mục của Biên bản giám định.

    Mỗi dòng: {sttN} {ten_phu_tung_N} {muc_do_N} và 4 ô đánh dấu 'x':
      {x_sua_chua_N} / {x_thay_the_N}            — cột "Đề xuất giải quyết"
      {x_thu_hoi_co_N} / {x_thu_hoi_khong_N}     — cột "Hiện vật thu hồi"

    Phương án "Sửa chữa" mặc định KHÔNG thu hồi.
    """
    for i in range(1, n + 1):
        pt = phu_tung[i - 1] if i - 1 < len(phu_tung) else None
        if pt is None:
            info[f'stt{i}'] = ''
            info[f'ten_phu_tung_{i}'] = ''
            info[f'muc_do_{i}'] = ''
            for col in ('x_sua_chua', 'x_thay_the', 'x_thu_hoi_co', 'x_thu_hoi_khong'):
                info[f'{col}_{i}'] = ''
            continue

        pa = pt['phuong_an'].lower()
        la_thay_the = 'thay thế' in pa
        co_thu_hoi = la_thay_the and 'thu hồi' in pa and 'không' not in pa

        info[f'stt{i}'] = str(pt.get('stt', i))
        info[f'ten_phu_tung_{i}'] = pt['ten']
        info[f'muc_do_{i}'] = pt.get('muc_do', '')
        info[f'x_sua_chua_{i}'] = '' if la_thay_the else 'x'
        info[f'x_thay_the_{i}'] = 'x' if la_thay_the else ''
        info[f'x_thu_hoi_co_{i}'] = 'x' if co_thu_hoi else ''
        info[f'x_thu_hoi_khong_{i}'] = '' if co_thu_hoi else 'x'


def phu_tung_thu_hoi(phu_tung):
    """Lọc phụ tùng phương án 'Thay thế có thu hồi' (loại 'không thu hồi')."""
    return [p for p in phu_tung
            if 'thu hồi' in p['phuong_an'].lower()
            and 'không' not in p['phuong_an'].lower()]


# ── Số tiền sang chữ ──────────────────────────────────────────────────
_CHU_SO = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín']
_DON_VI = ['', 'nghìn', 'triệu', 'tỷ']


def _doc_3_chu_so(n):
    tram, chuc, dv = n // 100, (n % 100) // 10, n % 10
    r = ''
    if tram > 0:
        r += _CHU_SO[tram] + ' trăm'
    if chuc == 0:
        if dv > 0 and tram > 0:
            r += ' lẻ ' + _CHU_SO[dv]
        elif dv > 0:
            r += _CHU_SO[dv]
    elif chuc == 1:
        r += ' mười'
        if dv == 5:
            r += ' lăm'
        elif dv > 0:
            r += ' ' + _CHU_SO[dv]
    else:
        r += ' ' + _CHU_SO[chuc] + ' mươi'
        if dv == 1:
            r += ' mốt'
        elif dv == 5:
            r += ' lăm'
        elif dv > 0:
            r += ' ' + _CHU_SO[dv]
    return r.strip()


def so_thanh_chu(so_str):
    """Chuyển số tiền VND sang chữ tiếng Việt.

    Theo cách đọc chuẩn: 15 → "mười lăm", 21 → "hai mươi mốt", 25 → "hai mươi lăm".
    Trả lại nguyên chuỗi đầu vào nếu không phải số.
    """
    try:
        so = int(str(so_str).replace(',', '').replace('.', '').replace(' ', '').strip())
    except (ValueError, AttributeError):
        return str(so_str)

    if so == 0:
        return 'Không đồng'

    groups = []
    tmp = so
    while tmp > 0:
        groups.append(tmp % 1000)
        tmp //= 1000

    parts = []
    for i in range(len(groups) - 1, -1, -1):
        if groups[i] > 0:
            txt = _doc_3_chu_so(groups[i])
            if _DON_VI[i]:
                txt += ' ' + _DON_VI[i]
            parts.append(txt)

    kq = ' '.join(parts).strip()
    return kq[0].upper() + kq[1:] + ' đồng'


# ── Xử lý XML trong .docx ─────────────────────────────────────────────
def merge_split_placeholders(xml):
    """Gộp các <w:r> liên tiếp trong cùng <w:p> mà nối text lại thành {placeholder}.

    Word hay tách `{ten_phu_tung_2}` thành `{ten_phu_tung_` + `2` + `}`; không gộp
    trước thì phép thay thế chuỗi sẽ trượt.
    """
    def process_para(para):
        # Chỉ match <w:t> hoặc <w:t ...>, KHÔNG match <w:tab> hay <w:tabs>
        t_re = re.compile(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)', re.DOTALL)
        matches = list(t_re.finditer(para))
        if not matches:
            return para

        texts = [m.group(2) for m in matches]
        combined = ''.join(texts)
        if '{' not in combined:
            return para

        char_map = []
        for idx, t in enumerate(texts):
            char_map.extend([idx] * len(t))

        placeholders_found = list(re.compile(r'\{[^{}]+\}').finditer(combined))
        if not placeholders_found:
            return para

        for ph in placeholders_found:
            s, e = ph.start(), ph.end() - 1
            if s >= len(char_map) or e >= len(char_map):
                continue
            si, ei = char_map[s], char_map[e]
            if si == ei:
                continue

            offset_in_si = s - sum(len(texts[k]) for k in range(si))
            offset_in_ei_end = e - sum(len(texts[k]) for k in range(ei)) + 1

            prefix = texts[si][:offset_in_si]
            suffix = texts[ei][offset_in_ei_end:]

            texts[si] = prefix + ph.group(0)
            texts[ei] = suffix
            for k in range(si + 1, ei):
                texts[k] = ''

        result = para
        for i in range(len(matches) - 1, -1, -1):
            m = matches[i]
            if ' ' in texts[i] and 'xml:space' not in m.group(1):
                new_t = '<w:t xml:space="preserve">' + texts[i] + m.group(3)
            else:
                new_t = m.group(1) + texts[i] + m.group(3)
            result = result[:m.start()] + new_t + result[m.end():]

        return result

    out = []
    pos = 0
    for pm in re.finditer(r'<w:p[ >]', xml):
        p_start = pm.start()
        p_end_idx = xml.find('</w:p>', p_start)
        if p_end_idx == -1:
            break
        p_end = p_end_idx + len('</w:p>')
        out.append(xml[pos:p_start])
        out.append(process_para(xml[p_start:p_end]))
        pos = p_end
    out.append(xml[pos:])
    return ''.join(out)


def apply_replacements(xml, data):
    for key, val in data.items():
        xml = xml.replace('{' + key + '}', str(val))
    return xml


def render(asset_name, output_name, info):
    """Đọc assets/<asset_name>, thay placeholder, ghi output/<output_name>.

    Trả về đường dẫn file đã tạo.
    """
    src = os.path.join(ASSETS_DIR, asset_name)
    dst = os.path.join(OUTPUT_DIR, output_name)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if _XML_PART_RE.match(item.filename):
                    xml = raw.decode('utf-8')
                    xml = merge_split_placeholders(xml)
                    xml = apply_replacements(xml, info)
                    _validate(xml, item.filename, output_name)
                    raw = xml.encode('utf-8')
                zout.writestr(item, raw)

    return dst


def _validate(xml, part, output_name):
    """XML vỡ thì Word không mở được file — chặn ngay tại đây."""
    import xml.etree.ElementTree as ET
    try:
        ET.fromstring(xml)
    except ET.ParseError as e:
        raise SystemExit(f'XML hỏng ở {part} của {output_name}: {e}')


def bao_cao_placeholder_con_sot(dst):
    """In các placeholder chưa được thay, để phát hiện template thiếu dữ liệu."""
    with zipfile.ZipFile(dst, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    remaining = re.findall(r'\{[a-zA-ZÀ-ỹ_][^{}]{1,50}\}', xml)
    remaining = [p for p in remaining if not p.startswith('{28A0')]
    if remaining:
        print('  Placeholder còn sót:', sorted(set(remaining)))
    else:
        print('  Tất cả placeholder đã được thay thế!')
