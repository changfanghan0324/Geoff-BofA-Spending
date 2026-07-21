#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_local_preview.py — 用桌面現有的年度 xlsx 產生 data.js（給你「先看到成品」）

它讀取這個資料夾裡所有 *_dad_card_professional_analysis.xlsx，
解析其中的交易明細，產生 data.js。
之後正式上線是靠 GitHub Action 去抓 Google Sheet（build_data.py），
這支只是讓你在連 Google Sheet 之前就能預覽網站長相。
"""

import glob
import os

import openpyxl

import build_data as B

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    files = sorted(glob.glob(os.path.join(ROOT, "*.xlsx")))
    if not files:
        print("找不到任何 .xlsx")
        return

    all_records = []
    for f in files:
        default_year = B.year_from_title(os.path.basename(f))
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        recs = B.parse_workbook(wb, default_year=default_year)
        wb.close()
        print("  %-55s -> %4d 筆" % (os.path.basename(f), len(recs)))
        all_records.extend(recs)

    config = B.load_config()
    payload = B.build_payload(all_records, config)
    out = B.write_data_js(payload)
    print("完成：%s（%d 個年份，%d 筆）" % (
        out, len(payload["years"]), sum(len(y["records"]) for y in payload["years"])))


if __name__ == "__main__":
    main()
