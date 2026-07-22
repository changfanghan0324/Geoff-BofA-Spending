#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — 把 Google Sheet（或本地 xlsx 活頁簿）解析成網站用的 data.js

用法（GitHub Action 會這樣呼叫）：
    python scripts/build_data.py

它會：
  1. 讀 config.json 拿到 sheet_id / title / currency
  2. 下載整份 Google 試算表（xlsx，包含所有年份分頁）
  3. 用「模糊欄位比對」解析每個分頁，不管欄名是中文還是英文都認得
  4. 產生 data.js（網站直接用 <script> 載入，不需要伺服器）

欄位對應（大小寫、中英文都可）：
  日期     ← Date / 日期
  種類     ← Category / 種類 / 類別 / 分類
  詳細說明 ← Description / 詳細 / 說明 / 摘要 / 項目
  花費價格 ← Expense Amount / 花費 / 支出 / 消費
  入金＆轉帳 ← Amount 為正的列 / 入金 / 轉帳 / 存入 / 收入
  餘額     ← Running Balance / 餘額 / 結餘
"""

import io
import json
import os
import re
import sys
import datetime
import urllib.request

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---- 欄位比對關鍵字（全部轉小寫後比對）-----------------------------------
KW = {
    "date":     (["日期", "date"],                              ["source", "display"]),
    "desc":     (["詳細", "說明", "摘要", "項目", "品項", "description", "desc", "item", "detail"], []),
    "category": (["種類", "類別", "分類", "category", "type"],    []),
    "expense":  (["花費", "支出", "消費", "expense", "spending", "debit"], []),
    "credit":   (["入金", "轉帳", "存入", "入帳", "收入", "credit", "deposit", "funding"], []),
    "balance":  (["餘額", "結餘", "balance"],                    []),
    "amount":   (["金額", "amount"],                            ["expense", "花費", "支出"]),
    "year":     (["年份", "年度", "year"],                      []),
}


# 分類正規化：把 2019~2021 的中文分類，對應到 2022~2026 使用的英文分類。
# 沒有現成對應的（餐飲、購物…）給一個新的英文正規名，前端會再翻成中文並標注僅 2019~2021。
CATEGORY_MAP = {
    # —— 直接併入既有分類 ——
    "教育": "Tuition & Education",
    "交通": "Travel & Transportation",
    "旅遊": "Travel & Transportation",
    "住房": "Housing & HOA",
    "手續費": "Fees & Banking",
    "轉帳入金": "Credits / Funding",
    "匯款入金": "Credits / Funding",
    "退款": "Credits / Funding",
    "開始": "Credits / Funding",          # 期初帳戶餘額
    # —— 保留細節：新增分類 ——
    "餐飲": "Dining & Food",
    "日常採買": "Groceries & Daily",
    "購物": "Shopping",
    "娛樂與訂閱": "Entertainment & Subscriptions",
    "個人照護": "Personal Care",
    "轉帳": "Personal Transfers",           # 個人轉帳（支出）
    "現金提領": "Cash Withdrawal",
    "郵務": "Postal & Shipping",
}


def normalize_category(cat):
    return CATEGORY_MAP.get(cat, cat)


def _norm(s):
    return re.sub(r"\s+", "", str(s).strip().lower()) if s is not None else ""


def find_col(headers, key):
    """回傳符合 key 的欄位索引，找不到回 None。"""
    include, exclude = KW[key]
    for i, h in enumerate(headers):
        hn = _norm(h)
        if not hn:
            continue
        if any(x in hn for x in exclude):
            continue
        if any(x in hn for x in include):
            return i
    return None


def to_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")   # 會計格式 (123) = 負數
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return None
    try:
        n = float(s)
        return -n if neg else n
    except ValueError:
        return None


def norm_date(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", s)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return s


def year_from_title(title):
    m = re.search(r"(20\d{2}|19\d{2})", str(title))
    return m.group(1) if m else None


def parse_sheet(ws, default_year=None):
    """把一個分頁解析成 record 陣列。認不出交易欄位就回空陣列。"""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # 在前 15 列裡找 header 列：必須有「日期」欄，且至少有金額/花費/餘額/描述其中之一
    header_idx = None
    headers = None
    for i, row in enumerate(rows[:15]):
        cells = list(row)
        if find_col(cells, "date") is not None and any(
            find_col(cells, k) is not None for k in ("amount", "expense", "balance", "desc", "credit")
        ):
            header_idx = i
            headers = cells
            break
    if header_idx is None:
        return []

    col = {k: find_col(headers, k) for k in KW}
    sheet_year = year_from_title(ws.title) or default_year

    def cell(row, key):
        idx = col[key]
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    records = []
    for row in rows[header_idx + 1:]:
        if row is None or all(c is None for c in row):
            continue

        date = norm_date(cell(row, "date"))
        exp = to_number(cell(row, "expense"))
        cr = to_number(cell(row, "credit"))
        bal = to_number(cell(row, "balance"))
        amt = to_number(cell(row, "amount"))
        desc = cell(row, "desc")
        desc = "" if desc is None else str(desc).strip()
        cat = cell(row, "category")
        cat = "" if cat is None else str(cat).strip()
        cat = normalize_category(cat)

        # 由帶正負的 Amount 推導 花費 / 入金
        if exp is None:
            exp = abs(amt) if (amt is not None and amt < 0) else 0.0
        else:
            exp = abs(exp)
        if cr is None:
            cr = amt if (amt is not None and amt > 0) else 0.0
        else:
            cr = abs(cr)
        amount_signed = amt if amt is not None else (cr - exp)

        # 跳過完全沒內容的雜訊列
        if not date and not desc and exp == 0 and cr == 0 and bal is None:
            continue
        # 跳過重複的表頭列
        if _norm(desc) in ("description", "詳細說明", "說明") and exp == 0 and cr == 0:
            continue

        yr = None
        yv = cell(row, "year")
        if yv is not None:
            yr = year_from_title(yv)
        yr = yr or sheet_year

        records.append({
            "date": date,
            "category": cat,
            "description": desc,
            "expense": round(exp, 2),
            "credit": round(cr, 2),
            "balance": None if bal is None else round(bal, 2),
            "amount": round(amount_signed, 2),
            "year": yr,
        })
    return records


def parse_workbook(wb, default_year=None):
    all_records = []
    for ws in wb.worksheets:
        all_records.extend(parse_sheet(ws, default_year=default_year))
    return all_records


def group_years(records):
    buckets = {}
    for r in records:
        yr = r.get("year") or (r["date"][:4] if r.get("date") else "未分類")
        buckets.setdefault(yr, []).append(r)

    years = []
    for yr in sorted(buckets.keys()):
        recs = buckets[yr]
        recs.sort(key=lambda r: (r["date"], r["description"]))

        total_expense = round(sum(r["expense"] for r in recs), 2)
        total_credit = round(sum(r["credit"] for r in recs), 2)
        end_balance = None
        for r in recs:
            if r["balance"] is not None:
                end_balance = r["balance"]

        # 分類統計（只看花費）
        cat_map = {}
        for r in recs:
            if r["expense"] > 0:
                c = r["category"] or "未分類"
                m = cat_map.setdefault(c, {"name": c, "total": 0.0, "count": 0})
                m["total"] += r["expense"]
                m["count"] += 1
        cats = sorted(cat_map.values(), key=lambda x: x["total"], reverse=True)
        for c in cats:
            c["total"] = round(c["total"], 2)
            c["share"] = round(c["total"] / total_expense, 4) if total_expense else 0

        # 每月收支（給趨勢圖）
        monthly = {}
        for r in recs:
            mkey = r["date"][:7] if len(r["date"]) >= 7 else "?"
            m = monthly.setdefault(mkey, {"month": mkey, "expense": 0.0, "credit": 0.0})
            m["expense"] += r["expense"]
            m["credit"] += r["credit"]
        months = sorted(monthly.values(), key=lambda x: x["month"])
        for m in months:
            m["expense"] = round(m["expense"], 2)
            m["credit"] = round(m["credit"], 2)

        # 清掉 record 內部用的 year 欄位，讓輸出精簡
        for r in recs:
            r.pop("year", None)

        years.append({
            "year": yr,
            "summary": {
                "total_expense": total_expense,
                "total_credit": total_credit,
                "net": round(total_credit - total_expense, 2),
                "count": len(recs),
                "end_balance": end_balance,
            },
            "categories": cats,
            "months": months,
            "records": recs,
        })
    return years


def load_config():
    path = os.path.join(ROOT, "config.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_payload(records, config):
    return {
        "generated_at": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "title": config.get("title", "消費紀錄"),
        "currency": config.get("currency_symbol", "$"),
        "years": group_years(records),
    }


def write_data_js(payload, path=None):
    path = path or os.path.join(ROOT, "data.js")
    js = "window.CARD_DATA = " + json.dumps(payload, ensure_ascii=False, indent=1) + ";\n"
    with open(path, "w", encoding="utf-8") as f:
        f.write(js)
    return path


def read_existing_payload(path):
    """讀取現有 data.js 裡的 JSON，讀不到回 None。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            t = f.read()
        return json.loads(t[t.index("{"):t.rindex("}") + 1])
    except Exception:
        return None


def same_data(a, b):
    """比較兩份 payload 是否實質相同（忽略 generated_at 時間戳）。"""
    if not a or not b:
        return False
    ax = {k: v for k, v in a.items() if k != "generated_at"}
    bx = {k: v for k, v in b.items() if k != "generated_at"}
    return ax == bx


def fetch_google_xlsx(sheet_id):
    url = "https://docs.google.com/spreadsheets/d/%s/export?format=xlsx" % sheet_id
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return io.BytesIO(resp.read())


def main():
    config = load_config()
    sheet_id = (config.get("sheet_id") or "").strip()
    if not sheet_id or sheet_id.startswith("PUT_YOUR"):
        print("!! 還沒設定 config.json 裡的 sheet_id，跳過抓取。", file=sys.stderr)
        sys.exit(1)

    print("下載 Google 試算表 ...")
    data = fetch_google_xlsx(sheet_id)
    wb = openpyxl.load_workbook(data, read_only=True, data_only=True)
    records = parse_workbook(wb)
    wb.close()
    if not records:
        print("!! 解析不到任何交易資料，請確認分頁欄位（日期/花費/餘額…）。", file=sys.stderr)
        sys.exit(1)

    payload = build_payload(records, config)

    # 只有資料真的變動才寫檔（忽略時間戳），這樣天天自動跑也不會產生無意義的 commit
    existing = read_existing_payload(os.path.join(ROOT, "data.js"))
    if same_data(existing, payload):
        print("資料沒有變動，data.js 保持不變（不會產生 commit）")
        return

    out = write_data_js(payload)
    print("完成：%s（%d 個年份，%d 筆紀錄）" % (
        out, len(payload["years"]), sum(len(y["records"]) for y in payload["years"])))


if __name__ == "__main__":
    main()
