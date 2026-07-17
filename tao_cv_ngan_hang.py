# -*- coding: utf-8 -*-
"""Mẫu G — Công văn gửi Ngân hàng (kèm Xác nhận của Ngân hàng/TCTD)."""
from pti_common import (load_data, so_thanh_chu, render,
                        bao_cao_placeholder_con_sot)

info, _ = load_data()

# Ước thiệt hại lấy từ tiền thanh toán trên màn hình hệ thống
so_tien = info.get('tien_tt', '0').replace(' ', '')
info['so_tien_thiet_hai_so'] = so_tien
info['so_tien_thiet_hai_chu'] = so_thanh_chu(so_tien)

dst = render('CV-Ngân hàng.docx', 'CV-Ngan-hang.docx', info)
print('Tạo xong:', dst)
print('Số tiền bằng chữ:', info['so_tien_thiet_hai_chu'])
bao_cao_placeholder_con_sot(dst)
