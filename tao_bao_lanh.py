# -*- coding: utf-8 -*-
"""Mẫu D — Bảo lãnh."""
from pti_common import load_data, render, bao_cao_placeholder_con_sot

info, _ = load_data()

dst = render('Bảo-lãnh.docx', 'Bao-lanh.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)
