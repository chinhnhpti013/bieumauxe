# -*- coding: utf-8 -*-
"""Mẫu I — Chấm dứt / Khôi phục bảo hiểm xe cơ giới PTI.

Cách dùng:
  1. Điền các trường NHẬP TAY trong phần "CẤU HÌNH NHẬP TAY" bên dưới
  2. Chạy: python tao_cham_dut_khoi_phuc.py
  3. File xuất: output/Cham-dut-khoi-phuc.docx

Nguồn dữ liệu tự động:
  - bien_so_xe, chu_xe, so_gcn_bh, phi_bh → đọc từ input/data.json
  - nam_hien_tai → tự tính từ ngày hệ thống

Nguồn dữ liệu nhập tay (phần bên dưới):
  - gcn_bh_ngay_cap: ngày cấp giấy chứng nhận BH
  - ngay/gio thời hạn BH VCX ô tô: cuối hạn BH (thường = gcn_bh_den_ngay)
  - Chứng từ doanh thu thực thu lần 1: ngày/tháng/năm và số tiền
  - Hạn thanh toán đợt 1: ngày/tháng/năm
"""
from datetime import datetime

from pti_common import load_data, render, bao_cao_placeholder_con_sot

# ════════════════════════════════════════════════════════════════════════
# CẤU HÌNH NHẬP TAY — Chỉnh sửa các giá trị bên dưới trước khi chạy
# ════════════════════════════════════════════════════════════════════════

NHAP_TAY = {
    # Ngày cấp GCN BH (DD/MM/YYYY) — từ GCNBH-VCX.pdf
    'gcn_bh_ngay_cap': '25/12/2025',

    # Thời hạn BH VCX ô tô — từ GCNBH-VCX.pdf: đến 09:25 ngày 28/12/2026
    'ngay_thoi_han_bao_hiem_vcx_oto': '28/12/2026',   # DD/MM/YYYY
    'gio_phut_thoi_han_bao_hiem_vcx_oto': '09:25',     # HH:MM

    # Giờ phút bắt đầu BH VCX — từ GCN BH
    'gcn_bh_gio_phut': '09:25',

    # Chứng từ doanh thu thực thu lần 1 — từ Phiếu XMP mục III
    # Dòng "Doanh thu Thực Thu" đầu tiên: ngày chứng từ 11/02/2026
    'ngay_chung_tu_doanh_thu_thuc_thu_lan_1':  '11',        # ngày
    'thang_chung_tu_doanh_thu_thuc_thu_lan_1': '02',        # tháng
    'nam_chung_tu_doanh_thu_thuc_thu_lan_1':   '2026',      # năm
    'tien_han_thanh_toan_dong_1': '18,551,250',              # số tiền đợt 1

    # Hạn thanh toán đợt 1 (kỳ 1) — từ Phiếu XMP: 24/01/2026
    'ngay_han_thanh_toan_dong_1':  '24',        # ngày
    'thang_han_thanh_toan_dong_1': '01',        # tháng
    'nam_han_thanh_toan_dong_1':   '2026',      # năm

    # Hạn nộp phí kỳ tiếp theo (kỳ 2) — từ Phiếu XMP: 24/02/2026
    'ke_tiep_ngay_nop_phi': '24/02/2026',
}

# ════════════════════════════════════════════════════════════════════════


def tong_hop_ngay(prefix, d):
    """Ghép thành chuỗi 'ngày DD tháng MM năm YYYY'."""
    ngay = d.get(f'ngay_{prefix}', '')
    thang = d.get(f'thang_{prefix}', '')
    nam = d.get(f'nam_{prefix}', '')
    return f'ngày {ngay} tháng {thang} năm {nam}' if ngay else ''


info, _ = load_data()

info['phi_bao_hiem_da_VAT'] = info.get('phi_bh', '')
info['nam_hien_tai'] = str(datetime.now().year)

info.update(NHAP_TAY)

info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1'] = tong_hop_ngay(
    'chung_tu_doanh_thu_thuc_thu_lan_1', NHAP_TAY)
info['ngay_thang_nam_han_thanh_toan_dong_1'] = tong_hop_ngay(
    'han_thanh_toan_dong_1', NHAP_TAY)

# Mẫu docx viết placeholder này có khoảng trắng bên trong: {gcn_bh_ ngay_cap}
info['gcn_bh_ ngay_cap'] = NHAP_TAY['gcn_bh_ngay_cap']

# Alias: mẫu dùng tên ngắn, không có hậu tố _lan_1 / _dong_1
for ngan, dai in [
    ('ngay_chung_tu_doanh_thu_thuc_thu',  'ngay_chung_tu_doanh_thu_thuc_thu_lan_1'),
    ('thang_chung_tu_doanh_thu_thuc_thu', 'thang_chung_tu_doanh_thu_thuc_thu_lan_1'),
    ('nam_chung_tu_doanh_thu_thuc_thu',   'nam_chung_tu_doanh_thu_thuc_thu_lan_1'),
    ('tien_han_thanh_toan',               'tien_han_thanh_toan_dong_1'),
    ('ngay_han_thanh_toan',               'ngay_han_thanh_toan_dong_1'),
    ('thang_han_thanh_toan',              'thang_han_thanh_toan_dong_1'),
    ('nam_han_thanh_toan',                'nam_han_thanh_toan_dong_1'),
]:
    info[ngan] = NHAP_TAY.get(dai, '')

info['ngay_thang_nam_han_thanh_toan'] = tong_hop_ngay('han_thanh_toan_dong_1', NHAP_TAY)

dst = render('Cham-dut-khoi-phuc.docx', 'Cham-dut-khoi-phuc.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)

print()
print('─── Dữ liệu đã điền ───────────────────────────────────')
print(f"  Biển số xe    : {info.get('bien_so_xe', '')}")
print(f"  Chủ xe        : {info.get('chu_xe', '')}")
print(f"  Số GCN BH     : {info.get('so_gcn_bh', '')}")
print(f"  Phí BH        : {info.get('phi_bao_hiem_da_VAT', '')} đồng")
print(f"  Ngày cấp GCN  : {NHAP_TAY['gcn_bh_ngay_cap']}")
print(f"  Thời hạn BH   : {NHAP_TAY['ngay_thoi_han_bao_hiem_vcx_oto']} "
      f"{NHAP_TAY['gio_phut_thoi_han_bao_hiem_vcx_oto']}")
print(f"  Chứng từ lần 1: {info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1']}")
print(f"  Tiền đợt 1    : {NHAP_TAY['tien_han_thanh_toan_dong_1']} đồng")
print(f"  Hạn TT đợt 1  : {info['ngay_thang_nam_han_thanh_toan_dong_1']}")
