# -*- coding: utf-8 -*-
"""Mẫu E — Thông báo tai nạn / Yêu cầu bồi thường."""
from pti_common import load_data, render, bao_cao_placeholder_con_sot

info, _ = load_data()

dst = render('TBTN-YCBT.docx', 'TBTN-YCBT.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)
