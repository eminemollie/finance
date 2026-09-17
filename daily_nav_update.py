#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
daily_nav_update.py
每日排程執行：抓取摩根多重收益基金最新淨值(MoneyDJ)與USD/TWD即期匯率(台灣銀行)，
更新「常駐加密Excel副本」(financial-workbook.enc) 裡的 NAV/匯率假設與淨值歷史列，
再重新產生 data.json 給手機版網頁讀取。

設計原則（跟 extract_data.py 的健檢哲學一致）：
- 抓取失敗、資料超出合理範圍 → 印警告、直接結束，不寫任何檔案，不讓排程失敗中斷。
- 只覆寫「NAV／匯率」這兩個原始輸入儲存格與稽核區塊，其餘公式/資料完全不動。
- 用 openpyxl 存檔前設定 fullCalcOnLoad，讓使用者之後用真正的 Excel 打開時，
  所有公式（貸款攤還、淨值等）會自動重新計算，不會停留在舊快取值。

執行方式：python3 daily_nav_update.py
需要環境變數 DASHBOARD_PASSWORD（跟 extract_data.py 共用同一組 GitHub Secret）。
"""
import os
import re
import io
import csv
import sys
import json
import datetime
import tempfile
import unicodedata

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import extract_data as ed

MONEYDJ_URL = 'https://www.moneydj.com/funddj/ya/yp010001.djhtm?a=jfzn3'  # JPM多重收益(美元對沖)-A股(穩定月配)
MONEYDJ_TECH_FUND_URL = 'https://www.moneydj.com/funddj/ya/yp010000.djhtm?a=acjf76'  # 摩根新興科技基金-一般型(新台幣)
BOT_CSV_URL = 'https://rate.bot.com.tw/xrt/flcsv/0/day'  # 台灣銀行牌告匯率CSV（輕量端點，主要來源）
FRANKFURTER_URL = 'https://api.frankfurter.dev/v2/rate/usd/twd'  # 歐洲央行每日參考匯率（公開API，備援來源）
TWSE_STOCK_DAY_ALL_URL = 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL'  # 證交所OpenAPI，全市場當日收盤價（主要來源）
TWSE_STOCK_DAY_ALL_CSV_URL = 'https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL?response=json'  # 證交所網站備援端點，實測回應其實是CSV格式（見_fetch_twse_stock_day_all_csv()註解）
ENC_PATH = 'financial-workbook.enc'
JSON_PATH = 'data.json'

NAV_MIN, NAV_MAX = 20.0, 200.0
FX_MIN, FX_MAX = 20.0, 45.0
TECH_FUND_NAV_MIN, TECH_FUND_NAV_MAX = 100.0, 2000.0  # 摩根新興科技基金淨值(新台幣計價)合理範圍
STOCK_PRICE_RANGES = {  # 個股/ETF收盤價合理範圍，抓到超出範圍的數字視為抓取異常（例如抓錯欄位）
    '0050': (30.0, 300.0),
    '0052': (15.0, 200.0),
}
STOCK_HOLDING_CODES = set(STOCK_PRICE_RANGES.keys())
TECH_FUND_NAME = '摩根新興科技證券投資信託基金'


def _fetch_moneydj_nav(url, nav_min, nav_max):
    """從 MoneyDJ 淨值表撈「淨值日期」欄與同一列裡的淨值欄位。抓不到／格式跑掉就丟例外。
    這裡不寫死欄位的完整標題文字——境外基金頁（如MONEYDJ_URL）欄位標題固定是「最新淨值」，
    但國內基金頁（如MONEYDJ_TECH_FUND_URL）欄位標題可能是「淨值」「淨值(最新)」等變體，
    只認「淨值日期」這個固定欄位，同一列裡再找下一個「文字裡有淨值兩個字」的欄位當作淨值本身，
    讓同一套解析邏輯可以共用給不同版型的MoneyDJ頁面。"""
    resp = requests.get(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }, timeout=20)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or 'utf-8'
    soup = BeautifulSoup(resp.text, 'html.parser')

    header_cells, data_cells, idx_nav = None, None, None
    for tr in soup.find_all('tr'):
        cells = [c.get_text(strip=True) for c in tr.find_all(['td', 'th'])]
        if '淨值日期' in cells:
            idx_date_candidate = cells.index('淨值日期')
            nav_candidates = [i for i, c in enumerate(cells) if i != idx_date_candidate and '淨值' in c]
            if not nav_candidates:
                continue
            header_cells = cells
            idx_nav = nav_candidates[0]
            nxt = tr.find_next_sibling('tr')
            if nxt:
                data_cells = [c.get_text(strip=True) for c in nxt.find_all(['td', 'th'])]
            break

    if not header_cells or not data_cells:
        raise RuntimeError('MoneyDJ 頁面結構可能已變動，找不到「淨值日期」表格')

    idx_date = header_cells.index('淨值日期')
    date_str = data_cells[idx_date].strip()
    nav_str = data_cells[idx_nav].strip()

    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', date_str)
    if not m:
        raise RuntimeError(f'淨值日期格式無法解析：{date_str!r}')
    nav_date = f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'

    try:
        nav = float(nav_str.replace(',', ''))
    except ValueError:
        raise RuntimeError(f'淨值無法轉成數字：{nav_str!r}')

    if not (nav_min <= nav <= nav_max):
        raise RuntimeError(f'抓到的淨值 {nav} 超出合理範圍({nav_min}~{nav_max})，可能抓錯欄位，放棄更新')

    return nav, nav_date


def fetch_nav_moneydj():
    """摩根多重收益基金(美元對沖)-A股(穩定月配)，美元計價，最新淨值。"""
    return _fetch_moneydj_nav(MONEYDJ_URL, NAV_MIN, NAV_MAX)


def fetch_nav_tech_fund():
    """摩根新興科技證券投資信託基金－一般型，新台幣計價，最新淨值。"""
    return _fetch_moneydj_nav(MONEYDJ_TECH_FUND_URL, TECH_FUND_NAV_MIN, TECH_FUND_NAV_MAX)


def _fetch_twse_stock_day_all_openapi():
    """證交所OpenAPI（主要來源）：乾淨的JSON陣列，每筆是{"Date":"1150916","Code":"0050",...,"ClosingPrice":"106.90",...}。
    2026-09-17實測：這個端點在GitHub Actions runner的網路環境下會回應不是合法JSON的內容
    （很可能是WAF/機器人驗證擋下，回應變成HTML挑戰頁，本地研究階段用WebFetch測試時是正常的，
    但WebFetch跟Actions runner的requests函式庫走的是不同網路路徑/IP，行為不保證一致——
    這個坑後來才發現，見下面_fetch_twse_stock_day_all_csv()這個備援來源）。
    回傳 {代號: (民國年日期字串, 收盤價字串)}。"""
    resp = requests.get(TWSE_STOCK_DAY_ALL_URL, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }, timeout=30)
    resp.raise_for_status()
    try:
        rows = resp.json()
    except ValueError:
        raise RuntimeError('證交所OpenAPI回應不是合法JSON，可能已改版或被擋下')
    if not isinstance(rows, list):
        raise RuntimeError('證交所OpenAPI回應格式不是預期的陣列，可能已改版')

    result = {}
    for row in rows:
        code = row.get('Code')
        if code is None:
            continue
        result[code] = (str(row.get('Date', '')), str(row.get('ClosingPrice', '')))
    return result


def _fetch_twse_stock_day_all_csv():
    """證交所網站備援端點：`www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL?response=json`。
    2026-09-17實測發現：雖然網址帶了`response=json`，這個端點實際回應的內容是**CSV格式**
    （不是JSON——response=json這個參數在這條路徑上沒有作用），欄位依序是
    日期,證券代號,證券名稱,成交股數,成交金額,開盤價,最高價,最低價,收盤價,漲跌價差,成交筆數，
    值用雙引號包起來，日期是民國年（例如"1150917"）。這跟主要來源(openapi.twse.com.tw)是不同的
    主機/路徑，即使主要來源被擋，這個備援端點也有機會正常回應（跟USD/TWD匯率的台灣銀行主頁
    vs CSV端點同一種「換一個端點繞過防護」的邏輯一致）。
    回傳 {代號: (民國年日期字串, 收盤價字串)}。"""
    resp = requests.get(TWSE_STOCK_DAY_ALL_CSV_URL, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or 'utf-8'
    text = resp.text.strip()
    if text.lower().startswith('<!doctype') or text.lower().startswith('<html'):
        raise RuntimeError('證交所備援CSV端點回應內容是HTML，可能被擋下')

    reader = csv.reader(io.StringIO(text))
    header = next(reader, None)
    if not header or '證券代號' not in header or '收盤價' not in header:
        raise RuntimeError('證交所備援CSV端點表頭跟預期不符，可能已改版')
    idx_date = header.index('日期')
    idx_code = header.index('證券代號')
    idx_close = header.index('收盤價')

    result = {}
    for row in reader:
        if len(row) <= max(idx_date, idx_code, idx_close):
            continue
        code = row[idx_code].strip()
        result[code] = (row[idx_date].strip(), row[idx_close].strip())
    return result


def fetch_twse_close_prices(codes):
    """抓全市場當日（或最近交易日）收盤價，篩出需要的代號，回傳 {代號: (收盤價, 日期ISO字串)}。
    2026-09-17新增備援機制：先試證交所OpenAPI（乾淨JSON），失敗（含被擋、回應非JSON）就自動
    改用證交所網站的CSV端點備援——跟既有 fetch_fx_bot() 的「主要來源失敗→自動換備援」設計
    原則一致。兩個來源都失敗、或篩出來的代號資料格式異常，才整批放棄（fail loud，不寫壞任何檔案）。"""
    raw = None
    try:
        raw = _fetch_twse_stock_day_all_openapi()
    except Exception as e:
        print(f'⚠️  證交所OpenAPI抓取失敗（{e}），改用備援來源(www.twse.com.tw CSV端點)')

    if raw is None:
        raw = _fetch_twse_stock_day_all_csv()

    found = {}
    for code in codes:
        if code not in raw:
            continue
        date_str, close_str = raw[code]
        m = re.match(r'(\d{3})(\d{2})(\d{2})', date_str)
        if not m:
            raise RuntimeError(f'{code} 的日期格式無法解析：{date_str!r}')
        date_iso = f'{int(m.group(1)) + 1911:04d}-{m.group(2)}-{m.group(3)}'
        try:
            close = float(close_str.replace(',', ''))
        except ValueError:
            raise RuntimeError(f'{code} 的收盤價無法轉成數字：{close_str!r}')
        lo, hi = STOCK_PRICE_RANGES.get(code, (1.0, 5000.0))
        if not (lo <= close <= hi):
            raise RuntimeError(f'{code} 抓到的收盤價 {close} 超出合理範圍({lo}~{hi})，可能抓錯欄位')
        found[code] = (close, date_iso)

    missing = set(codes) - found.keys()
    if missing:
        raise RuntimeError(f'證交所回應中找不到以下代號的資料：{sorted(missing)}')
    return found


def _fetch_fx_bot_csv():
    """從台灣銀行牌告匯率CSV端點抓 USD 即期匯率買入/賣出，回傳兩者中價。
    台灣銀行的主要匯率頁面(rate.bot.com.tw/xrt)有機器人驗證(Radware)保護，
    一般HTTP請求（無法執行JS）會被擋下、回傳一個驗證挑戰頁面而不是真正資料，
    這裡改用它另外提供的CSV下載端點，一般不會經過同一層驗證。"""
    resp = requests.get(BOT_CSV_URL, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }, timeout=20)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or 'utf-8'
    text = resp.text.strip()

    if text.lower().startswith('<!doctype') or text.lower().startswith('<html'):
        raise RuntimeError('回應內容是HTML而非CSV，可能被機器人驗證擋下')

    reader = csv.reader(io.StringIO(text))
    for row in reader:
        cells = [c.strip() for c in row]
        if not cells or cells[0] != 'USD':
            continue
        nums = []
        for v in cells[1:]:
            try:
                fv = float(v)
                if fv > 0:
                    nums.append(fv)
            except ValueError:
                continue
        # 欄位順序一般是：現金買入, 現金賣出, 即期買入, 即期賣出（取最後兩個當即期匯率）
        if len(nums) >= 2:
            buy, sell = nums[-2], nums[-1]
            fx = round((buy + sell) / 2, 4)
            if FX_MIN <= fx <= FX_MAX:
                return fx
        raise RuntimeError(f'USD這一列數字格式跟預期不符：{cells}')

    raise RuntimeError('CSV裡找不到USD這一列，格式可能已變動')


def _fetch_fx_frankfurter():
    """備援匯率來源：Frankfurter（歐洲央行每日參考匯率），公開API，不會有機器人驗證問題。
    注意：Frankfurter在2026年改版過API路徑，舊版 v1/latest?base=...&symbols=... 已回404，
    現在的端點是 v2/rate/{base}/{quote}，回應格式也不同（單一 rate 欄位，不是 rates 物件）。"""
    resp = requests.get(FRANKFURTER_URL, timeout=20)
    resp.raise_for_status()
    payload = resp.json()
    fx = payload.get('rate')
    if fx is None:
        raise RuntimeError(f'回應中找不到匯率欄位：{payload}')
    fx = round(float(fx), 4)
    if not (FX_MIN <= fx <= FX_MAX):
        raise RuntimeError(f'抓到的匯率 {fx} 超出合理範圍({FX_MIN}~{FX_MAX})')
    return fx


def fetch_fx_bot():
    """抓 USD/TWD 匯率。優先用台灣銀行牌告匯率CSV，失敗（含被機器人驗證擋下）就改用歐洲央行參考匯率當備援。
    回傳 (匯率, 資料來源說明文字)。"""
    try:
        fx = _fetch_fx_bot_csv()
        return fx, '台灣銀行牌告匯率'
    except Exception as e:
        print(f'⚠️  台灣銀行匯率來源抓取失敗（{e}），改用備援來源(Frankfurter/歐洲央行參考匯率)')

    fx = _fetch_fx_frankfurter()
    return fx, 'Frankfurter歐洲央行參考匯率(備援)'


def months_between(start, today):
    if start is None:
        return 0
    months = (today.year - start.year) * 12 + (today.month - start.month)
    if today.day < start.day:
        months -= 1
    return max(0, months)


def remaining_loan_balance(principal, annual_rate, years, start_iso, today):
    if not start_iso:
        return principal
    start = datetime.date.fromisoformat(start_iso)
    n = int(round(years * 12))
    p = min(months_between(start, today), n)
    r = annual_rate / 12
    if r == 0:
        return max(0.0, principal * (1 - p / n)) if n else principal
    num = (1 + r) ** n - (1 + r) ** p
    den = (1 + r) ** n - 1
    return principal * num / den if den else principal


def remaining_insurance(premium, years, start_iso, today):
    if not start_iso or not years:
        return premium
    start = datetime.date.fromisoformat(start_iso)
    n = int(round(years * 12))
    p = min(months_between(start, today), n)
    return max(0.0, premium * (1 - p / n)) if n else premium


def compute_snapshot(data, today):
    fund_mv = data['totalUnits'] * data['nav'] * data['fx']
    a = data['assetsDefault']
    other_assets = a.get('north_re', 0) + a.get('east_re', 0) + a.get('stocks', 0) + a.get('cash', 0)
    total_assets = fund_mv + other_assets

    total_debt = 0.0
    for loan in data['loans']:
        total_debt += remaining_loan_balance(loan['principal'], loan['rate'], loan['years'], loan['start'], today)
    for ins in data['insurance']:
        total_debt += remaining_insurance(ins['premium'], ins['years'], ins['start'], today)

    net_worth = total_assets - total_debt
    return round(fund_mv), round(total_assets), round(total_debt), round(net_worth)


def update_assumption_cells(ws_bs, nav, fx, updated_at_str, fx_source):
    """更新資產負債表『市值假設』與『自動更新狀態』區塊。用label比對儲存格位置，不依賴寫死的row編號。"""
    found_nav = found_fx = found_audit = 0
    for r in range(1, ws_bs.max_row + 1):
        label = ws_bs.cell(row=r, column=2).value
        if label == '基金最新淨值(NAV,USD)':
            ws_bs.cell(row=r, column=3).value = nav
            found_nav += 1
        elif label == '目前匯率(USD/TWD)':
            ws_bs.cell(row=r, column=3).value = fx
            found_fx += 1
        elif label == '最後自動更新時間':
            ws_bs.cell(row=r, column=3).value = updated_at_str
            found_audit += 1
        elif label == '資料來源':
            ws_bs.cell(row=r, column=3).value = f'MoneyDJ（基金淨值）／{fx_source}（USD/TWD匯率）'
        elif label == '更新方式':
            ws_bs.cell(row=r, column=3).value = '每日自動（也可手動覆寫上方 NAV／匯率）'
    if found_nav == 0 or found_fx == 0:
        raise RuntimeError('資產負債表找不到 NAV／匯率假設儲存格，Excel版面可能被改動過，放棄更新以免寫錯位置')
    if found_audit == 0:
        print('⚠️  找不到「自動更新狀態」稽核區塊（可能是舊版檔案），僅更新NAV/匯率，不寫入稽核欄位')


def update_stock_positions(ws_pos, stock_prices, tech_fund_nav, tech_fund_nav_date):
    """更新「股票部位」分頁每一列的最新價格／市值／未實現損益／價格日期／資料來源，
    回傳這次算出來的總市值（round過的整數，寫回資產負債表「股票市值」用）。
    用A欄的代號（0050/0052/摩根新興科技證券投資信託基金）比對儲存格位置，不依賴寫死的row編號；
    市值／未實現損益直接用Python算好寫入literal數字（不依賴Excel公式重算），
    跟整份系統既有的「基金市值」算法（compute_snapshot）走同一套設計原則。"""
    total_mv = 0.0
    found = 0
    r = 5
    while ws_pos.cell(row=r, column=1).value is not None:
        code = str(ws_pos.cell(row=r, column=1).value).strip()
        qty = ws_pos.cell(row=r, column=3).value or 0
        cost = ws_pos.cell(row=r, column=4).value or 0

        if code in stock_prices:
            price, price_date = stock_prices[code]
            source = '證交所（TWSE）每日收盤價'
        elif code == TECH_FUND_NAME:
            price, price_date = tech_fund_nav, tech_fund_nav_date
            source = 'MoneyDJ 基金淨值'
        else:
            r += 1
            continue

        mv = qty * price
        pnl = (price - cost) * qty

        ws_pos.cell(row=r, column=5).value = price
        ws_pos.cell(row=r, column=6).value = round(mv)
        ws_pos.cell(row=r, column=7).value = round(pnl)
        ws_pos.cell(row=r, column=8).value = price_date
        ws_pos.cell(row=r, column=9).value = f'{source}（每日自動）'

        total_mv += mv
        found += 1
        r += 1

    if found == 0:
        raise RuntimeError('「股票部位」分頁找不到任何可更新的持股列（比對代號0050/0052/摩根新興科技證券投資信託基金皆失敗），版面可能被改動過，放棄更新以免寫錯位置')

    return round(total_mv)


def update_stock_market_value_cell(ws_bs, stock_total):
    """把「股票部位」分頁算出來的總市值寫回資產負債表的「股票市值」列，
    從手動輸入改成每日自動計算。用label比對儲存格位置，不依賴寫死的row編號。"""
    # 用 startswith 而非「文字裡有出現」比對：資產負債表下方③的使用說明文字裡剛好也提到「股票市值」
    # 這幾個字，但那一列是B~D合併儲存格（純說明文字），用「in」比對會誤命中、寫入合併儲存格觸發
    # openpyxl的MergedCell唯讀例外；真正的目標列標籤固定是以「股票市值」開頭。
    found = 0
    for r in range(1, ws_bs.max_row + 1):
        label = ws_bs.cell(row=r, column=2).value
        if label and str(label).startswith('股票市值'):
            ws_bs.cell(row=r, column=3).value = stock_total
            ws_bs.cell(row=r, column=4).value = '＝0050/0052/摩根新興科技基金市值加總（每日自動更新，明細見「股票部位」分頁）'
            found += 1
    if found == 0:
        raise RuntimeError('資產負債表找不到「股票市值」儲存格，Excel版面可能被改動過，放棄更新以免寫錯位置')


def upsert_nav_history_row(ws_hist, today_iso, nav, fx, fund_mv, total_assets, total_debt, net_worth, source):
    """把今天的快照寫入淨值歷史分頁；同一天重複執行則更新原列，不重複新增。"""
    target_row = None
    r = 5
    while ws_hist.cell(row=r, column=1).value is not None:
        v = ws_hist.cell(row=r, column=1).value
        v_iso = v.date().isoformat() if isinstance(v, datetime.datetime) else (v.isoformat() if isinstance(v, datetime.date) else None)
        if v_iso == today_iso:
            target_row = r
            break
        r += 1
    if target_row is None:
        target_row = r  # 第一個空白列

    date_val = datetime.date.fromisoformat(today_iso)
    values = [date_val, nav, fx, fund_mv, total_assets, total_debt, net_worth, source]
    for i, v in enumerate(values, start=1):
        c = ws_hist.cell(row=target_row, column=i, value=v)
        if i == 1:
            c.number_format = 'yyyy/mm/dd'
        elif i in (4, 5, 6, 7):
            c.number_format = '#,##0'
        elif i in (2, 3):
            c.number_format = '0.00'


# ────────────────────────────────────────────────────────────────
# 2026-09-17新增：整份Excel共用的色彩／字體慣例（比照儀表板分頁第29~33列
# 「使用說明／色彩圖例」實際記載的規則，逐一比對既有分頁核對出來的既定風格）：
#   標題列(第1列)：深藍底(FF1F3864)、白色粗體、Noto Sans CJK SC 15號
#   區塊標題列(如「各項貸款明細」)：中藍底(FF2E75B6)、白色粗體、11號
#   表頭列(欄位名稱那一列)：淺藍底(FFD9E2F3)、黑色粗體
#   資料格：藍字(FF0000FF)＝原始輸入的事實（金額/日期/利率等）；
#           黑字(FF000000)＝同分頁自動計算公式；
#           綠字(FF008000)＝跨分頁連結公式（不要手動修改）
#   關鍵假設／需要定期更新的欄位：黃底(FFFFFF00)
# 這裡定義成常數，供下面 restyle_stock_positions_sheet()／autofit_other_sheets()
# 共用，避免顏色代碼分散寫死在各處。
# ────────────────────────────────────────────────────────────────
TITLE_FILL = PatternFill(fill_type='solid', fgColor='FF1F3864')
SECTION_FILL = PatternFill(fill_type='solid', fgColor='FF2E75B6')
HEADER_FILL = PatternFill(fill_type='solid', fgColor='FFD9E2F3')
TITLE_FONT = Font(name='Noto Sans CJK SC', size=15, bold=True, color='FFFFFFFF')
HEADER_FONT = Font(name='Noto Sans CJK SC', size=9, bold=True, color='FF000000')
BLUE_INPUT_COLOR = 'FF0000FF'   # 原始輸入事實
BLACK_FORMULA_COLOR = 'FF000000'  # 同分頁公式
LOCKED_LAYOUT_SHEETS = ('淨值歷史', '儀表板')  # 已由Norris手動鎖定固定版面，不要被其他版面調整邏輯動到


def restyle_stock_positions_sheet(ws):
    """把「股票部位」分頁的顏色／字體改成跟其他既有分頁完全一致的風格
    （2026-09-17發現這個分頁是新增的，建立時沒有比照其他分頁的既定配色慣例，
    標題沒有底色、表頭用了不同的藍色且是白字、資料格是預設Calibri字體，
    在整份Excel裡顯得不統一，Norris反應「各頁面風格我希望都統一一致」）。
    每天都會重新套用一次，寫法上是「設成固定值」，不是疊加，所以重複執行不會壞掉。"""
    # 標題列：深藍底、白字粗體（比照其他分頁）
    title_cell = ws.cell(row=1, column=1)
    title_cell.fill = TITLE_FILL
    title_cell.font = TITLE_FONT

    # 表頭列（第4列）：淺藍底、黑字粗體（原本誤用了跟其他分頁不同的藍底白字）
    for c in range(1, 10):
        cell = ws.cell(row=4, column=c)
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT

    # 資料列（第5列起）：比照儀表板色彩圖例——
    # 持有數量／均價／最新價格／價格日期／資料來源＝原始事實(藍字)；
    # 市值／未實現損益＝同分頁公式(黑字)；代號／名稱＝標籤文字(黑字，不特別上色)
    r = 5
    while ws.cell(row=r, column=1).value is not None:
        for col in (1, 2):  # 代號、名稱：標籤文字
            cell = ws.cell(row=r, column=col)
            cell.font = Font(name='Noto Sans CJK SC', size=11, color=BLACK_FORMULA_COLOR)
        for col in (3, 4, 5, 8, 9):  # 持有數量／均價／最新價格／價格日期／資料來源：原始事實
            cell = ws.cell(row=r, column=col)
            is_text = isinstance(cell.value, str)
            cell.font = Font(name='Noto Sans CJK SC' if is_text else 'Arial', size=11, color=BLUE_INPUT_COLOR)
        for col in (6, 7):  # 市值／未實現損益：同分頁公式
            cell = ws.cell(row=r, column=col)
            cell.font = Font(name='Arial', size=11, color=BLACK_FORMULA_COLOR)
        r += 1

    # 小計列：維持粗體，顏色改回黑字（公式）
    total_row = r
    for col in (6, 7):
        cell = ws.cell(row=total_row, column=col)
        if cell.value is not None:
            cell.font = Font(name='Arial', size=11, bold=True, color=BLACK_FORMULA_COLOR)
    label_cell = ws.cell(row=total_row, column=2)
    if label_cell.value is not None:
        label_cell.font = Font(name='Noto Sans CJK SC', size=11, bold=True, color=BLACK_FORMULA_COLOR)


def _display_width(text):
    """粗略估計字串在Excel等寬字體下的顯示寬度（中日韓字元／全形符號算2個單位，
    半形字元算1個單位），只是用來抓「欄寬夠不夠」的概算，不是精確排版計算。"""
    w = 0.0
    for ch in str(text):
        ea = unicodedata.east_asian_width(ch)
        if ea in ('W', 'F'):
            w += 2.0
        elif ea == 'A':
            w += 1.7
        else:
            w += 1.0
    return w


def _cell_display_text(cell):
    """回傳儲存格「大概會顯示成什麼文字」，用來估計需要的欄寬。
    公式儲存格沒辦法用openpyxl知道實際計算結果，一律跳過不納入估計
    （既有欄寬已經是照公式實際跑出來的結果調過，不太可能因為這個誤判過窄）。"""
    v = cell.value
    if v is None:
        return None
    if isinstance(v, str):
        if v.startswith('='):
            return None
        return v
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        nf = cell.number_format or ''
        if '.00' in nf:
            return f'{v:,.2f}' if ',' in nf else f'{v:.2f}'
        if ',' in nf or '#' in nf:
            return f'{v:,.0f}'
        return str(v)
    if isinstance(v, (datetime.date, datetime.datetime)):
        return '2026/09/16'  # 固定用這個當日期顯示寬度的估計基準，跟既有日期格式一致
    return str(v)


def autofit_other_sheets(wb, max_col_width=45, min_width=6, padding=2, wrap_row_height=32):
    """2026-09-17新增：Norris反應「各個頁面的字體大小都幫我自動設定到格子內」——
    openpyxl沒辦法真的量文字算圖元寬度，也不適合自動縮小字體（會讓整份Excel字體
    大小變得亂七八糟、更不統一），所以改成「自動加寬欄位」達到同樣的效果：
    掃過每個分頁每一格目前的內容，抓出每一欄實際需要的寬度，欄寬不夠就加寬到
    剛好放得下（只加寬、不縮小，不會破壞任何人手動調過的比例）。

    「淨值歷史」「儀表板」已經是Norris手動鎖定的固定版面（見repair_layout_and_charts
    的說明），這裡明確跳過，不去動它們的欄寬。有設定自動換行(wrap_text)的儲存格
    也跳過不列入欄寬估計——那些儲存格本來就是設計成「欄寬固定、用換行＋列高裝下
    較長文字」，不應該被這裡的邏輯拉寬。公式儲存格因為openpyxl讀不到實際計算結果，
    同樣跳過不列入估計。**跨欄合併儲存格（例如標題列A1:M1、長說明文字合併成一整列）
    也整個跳過**——這種儲存格的文字本來就是設計成橫跨好幾欄顯示，如果照單一欄位去
    估計需要的寬度，會把全部文字的寬度誤判成只有第一欄要承受，導致那一欄被離譜地
    拉超寬（這是第一版實作時實際測試中踩到的坑，特別記錄下來避免以後重犯）。

    **單一儲存格內容長度超過 max_col_width 時，不會把整欄硬撐到那麼寬**（這樣同一欄
    裡其他列的短內容會被拉得不成比例）——改成只針對那一個儲存格開自動換行＋加高
    列高，欄寬維持在合理上限，做法跟「淨值歷史」H欄／儀表板「資料來源」那組已經在
    用的「固定欄寬＋wrap＋加高列高」模式一致。

    每天都會重新跑一次，只會讓欄位越來越寬／列越來越高（或維持不變），不會產生
    累積性的錯誤。"""
    for sheet_name in wb.sheetnames:
        if sheet_name in LOCKED_LAYOUT_SHEETS:
            continue
        ws = wb[sheet_name]
        merged_coords = set()
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col != merged_range.max_col:  # 跨欄才需要排除；同欄多列合併不影響欄寬估計
                for row_cells in ws[merged_range.coord]:
                    for c in row_cells:
                        merged_coords.add(c.coordinate)

        col_max = {}
        overflow_cells = []  # 單一儲存格本身就超過欄寬上限，改用wrap+加高列高處理，不列入欄寬估計
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if cell.coordinate in merged_coords:
                    continue
                if cell.alignment and cell.alignment.wrap_text:
                    continue
                text = _cell_display_text(cell)
                if text is None:
                    continue
                needed = _display_width(text)
                if needed > max_col_width:
                    overflow_cells.append(cell)
                    continue
                col_letter = cell.column_letter
                col_max[col_letter] = max(col_max.get(col_letter, 0.0), needed)

        for col_letter, needed in col_max.items():
            needed_width = min(max_col_width, max(min_width, needed + padding))
            current = ws.column_dimensions[col_letter].width
            if current is None or needed_width > current:
                ws.column_dimensions[col_letter].width = round(needed_width, 1)

        for cell in overflow_cells:
            old_align = cell.alignment
            cell.alignment = Alignment(
                wrap_text=True,
                horizontal=old_align.horizontal if old_align else None,
                vertical=(old_align.vertical if old_align and old_align.vertical else 'top'),
            )
            current_h = ws.row_dimensions[cell.row].height
            if current_h is None or current_h < wrap_row_height:
                ws.row_dimensions[cell.row].height = wrap_row_height


def repair_layout_and_charts(wb):
    """修正「淨值歷史」「儀表板」兩個分頁的版面問題，並移除淨值走勢圖表。

    2026-09-17：Norris手動在Excel裡把文字大小／欄寬／對齊都調整成他滿意的樣子後，
    明確要求「之後的文字大小及位置就照這版本運行」——所以這裡不再是「只加寬/加高、
    不縮小」的最小值邏輯，而是每天都把版面**設成跟Norris那份定案版本完全一樣的固定值**
    （欄寬、列高、字體大小、對齊方式），不管前一天是什麼狀態，每天執行都會被重設成同一組
    數字，確保每日自動更新不會不小心把Norris手動調過的版面又蓋回去。

    定案版本的設定（讀Norris上傳的財務管理系統.xlsx比對確認）：
    1) A2說明文字：Noto Sans CJK SC、10號、斜體、灰色(FF5B6472)，自動換行，列高34.05
    2) 淨值歷史欄寬固定為 A12/B14/C13/D18/E13/F13/G16/H36
    3) 「更新來源」（H欄）：每一列已有資料的儲存格置中對齊＋自動換行＋字體縮小到Calibri 9號、
       列高30，避免長度不固定的說明文字整段溢出格子
    4) 直接移除「淨值歷史」「儀表板」兩個工作表的淨值走勢圖表（不管原本是折線圖還是柱狀圖，
       Norris反應圖表看不懂、乾脆整個拿掉，見對話記錄）
    5) 儀表板欄寬固定為 C22/D24/E22；「🔄 自動更新狀態」區塊「最後自動更新時間」「資料來源」
       兩列的值儲存格改成自動換行＋垂直靠上對齊、列高30；「資料來源」字體額外縮小到Arial 10號
       （保留原本的綠字FF008000，只改字體大小），因為這欄內容長度不固定，比「最後自動更新
       時間」更容易被裁掉
    6) 清掉儀表板上殘留的「📈 淨值走勢（近期...）」標題文字——圖表已經拿掉了，這行介紹圖表的
       標題文字留著沒有意義，Norris要求直接拿掉
    這個函式每天都會執行一次，刻意設計成重複執行也不會壞掉或疊加錯誤
    （每次都是「設成同樣的版面設定」，不是疊加或累積修改）。"""
    ws_hist = wb['淨值歷史']

    # 1) A2 說明文字：字體、對齊、列高都照Norris調整過的版本固定設定
    a2 = ws_hist['A2']
    a2.font = Font(name='Noto Sans CJK SC', size=10, italic=True, color='FF5B6472')
    a2.alignment = Alignment(wrap_text=True, vertical='center')
    ws_hist.row_dimensions[2].height = 34.05

    # 淨值歷史欄寬：固定設成Norris定案版本的數字（不是最小值，每天都重設成同一組）
    for col_letter, width in (('A', 12), ('B', 14), ('C', 13), ('D', 18), ('E', 13), ('F', 13), ('G', 16), ('H', 36)):
        ws_hist.column_dimensions[col_letter].width = width

    # 2) H欄（更新來源）每一列已經有資料的儲存格：置中對齊＋自動換行＋縮小字體＋固定列高
    r = 5
    while ws_hist.cell(row=r, column=1).value is not None:
        h_cell = ws_hist.cell(row=r, column=8)
        h_cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        h_cell.font = Font(name='Calibri', size=9)
        ws_hist.row_dimensions[r].height = 30
        r += 1

    # 3) 移除淨值走勢圖表（不管原本是折線圖還是柱狀圖）
    for sheet_name in ('淨值歷史', '儀表板'):
        wb[sheet_name]._charts = []

    # 4) 儀表板「自動更新狀態」區塊：欄寬固定＋對齊與字體照Norris定案版本
    if '儀表板' in wb.sheetnames:
        dash = wb['儀表板']
        for col_letter, width in (('C', 22), ('D', 24), ('E', 22)):
            dash.column_dimensions[col_letter].width = width

        for r in range(1, dash.max_row + 1):
            label = dash.cell(row=r, column=2).value
            if label == '最後自動更新時間':
                value_cell = dash.cell(row=r, column=3)
                value_cell.alignment = Alignment(wrap_text=True, vertical='top')
                dash.row_dimensions[r].height = 30
            elif label == '資料來源':
                value_cell = dash.cell(row=r, column=3)
                value_cell.alignment = Alignment(wrap_text=True, vertical='top')
                value_cell.font = Font(name='Arial', size=10, color='FF008000')
                dash.row_dimensions[r].height = 30

        # 5) 移除殘留的「📈 淨值走勢」標題文字（圖表已移除，這行文字沒有意義了）
        banner_text = '📈 淨值走勢（近期，資料來自「淨值歷史」分頁）'
        for r in range(1, dash.max_row + 1):
            if dash.cell(row=r, column=1).value == banner_text:
                for merged_range in list(dash.merged_cells.ranges):
                    if merged_range.min_row == r and merged_range.max_row == r:
                        dash.unmerge_cells(str(merged_range))
                for c in range(1, dash.max_column + 1):
                    cell = dash.cell(row=r, column=c)
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.font = Font()
                break


def main():
    # 2026-09-17安全性/穩健性修正：這個函式原本不管遇到什麼狀況都 return 0（成功），
    # 包括NAV/匯率抓取失敗、Excel結構跑掉寫不進去等「真正需要Norris回來處理」的情況——
    # 導致GitHub Actions畫面永遠顯示綠勾勾，Norris不會收到GitHub內建的失敗通知信，
    # 淨值可能已經好幾天沒真的更新了卻完全不會發現。現在改成：真正的執行失敗（抓取失敗、
    # 寫入失敗）回傳非0，讓這次執行在Actions畫面上顯示成失敗、觸發失敗通知信；只有「還沒
    # 設定好、本來就預期會這樣」的情況（例如DASHBOARD_PASSWORD還沒設定、常駐Excel副本
    # 還沒建立）維持return 0，不當作錯誤。
    password = os.environ.get('DASHBOARD_PASSWORD')
    if not password:
        print('[錯誤] 未設定 DASHBOARD_PASSWORD，無法解密常駐Excel副本，結束')
        return 1

    if not os.path.exists(ENC_PATH):
        print(f'⚠️  找不到 {ENC_PATH}（可能還沒手動 push 過一次新版Excel建立常駐副本），結束（尚未設定好，不視為錯誤）')
        return 0

    # ── 1) 抓取最新 NAV / 匯率，任何一個失敗就整個放棄，不動任何檔案 ──
    try:
        nav, nav_date = fetch_nav_moneydj()
        print(f'✅ MoneyDJ 淨值：{nav}（淨值日期 {nav_date}）')
    except Exception as e:
        print(f'[錯誤] 抓取基金淨值失敗，略過本次自動更新：{e}')
        return 1

    try:
        fx, fx_source = fetch_fx_bot()
        print(f'✅ USD/TWD 匯率：{fx}（來源：{fx_source}）')
    except Exception as e:
        print(f'[錯誤] 抓取匯率失敗（含備援來源），略過本次自動更新：{e}')
        return 1

    # 2026-09-17新增：抓取本人持有的0050/0052收盤價與摩根新興科技基金淨值，
    # 跟上面NAV/匯率一樣，失敗就整個放棄、不動任何檔案（fail loud，不要悄悄用舊資料）
    try:
        stock_prices = fetch_twse_close_prices(STOCK_HOLDING_CODES)
        for code in sorted(stock_prices):
            price, price_date = stock_prices[code]
            print(f'✅ {code} 收盤價：{price}（{price_date}）')
    except Exception as e:
        print(f'[錯誤] 抓取0050/0052收盤價失敗，略過本次自動更新：{e}')
        return 1

    try:
        tech_fund_nav, tech_fund_nav_date = fetch_nav_tech_fund()
        print(f'✅ 摩根新興科技基金淨值：{tech_fund_nav}（淨值日期 {tech_fund_nav_date}）')
    except Exception as e:
        print(f'[錯誤] 抓取摩根新興科技基金淨值失敗，略過本次自動更新：{e}')
        return 1

    now = datetime.datetime.now()
    today_iso = now.date().isoformat()
    updated_at_str = now.strftime('%Y-%m-%d %H:%M')

    # ── 2) 解密常駐 Excel 副本 ──
    with open(ENC_PATH, encoding='utf-8') as f:
        enc_raw = json.load(f)
    xlsx_bytes = ed.decrypt_bytes(enc_raw['data'], password)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_xlsx = os.path.join(tmpdir, 'financial-workbook.xlsx')
        with open(tmp_xlsx, 'wb') as f:
            f.write(xlsx_bytes)

        # ── 3) 更新 NAV／匯率／稽核欄位，以及本人股票／基金部位市值（保留所有公式，不動其他任何資料）──
        wb = openpyxl.load_workbook(tmp_xlsx)
        wb.calculation.fullCalcOnLoad = True
        try:
            update_assumption_cells(wb['資產負債表'], nav, fx, updated_at_str, fx_source)
            if '股票部位' in wb.sheetnames:
                stock_total = update_stock_positions(wb['股票部位'], stock_prices, tech_fund_nav, tech_fund_nav_date)
                update_stock_market_value_cell(wb['資產負債表'], stock_total)
                print(f'✅ 股票／基金部位總市值：{stock_total:,}')
            else:
                print('⚠️  找不到「股票部位」分頁（可能還沒上傳新版Excel建立），本次跳過股票市值自動更新，維持目前的手動輸入值')
        except Exception as e:
            print(f'[錯誤] {e}')
            return 1
        wb.save(tmp_xlsx)

        # ── 4) 用 extract_data 的既有邏輯重新讀取一次，取得最新單位數／資產／負債結構 ──
        data = ed.extract(tmp_xlsx)
        fund_mv, total_assets, total_debt, net_worth = compute_snapshot(data, now.date())
        print(f'📊 快照：基金市值 {fund_mv:,} / 總資產 {total_assets:,} / 總負債 {total_debt:,} / 淨值 {net_worth:,}')

        # ── 5) 寫入／更新今天的淨值歷史列 ──
        wb2 = openpyxl.load_workbook(tmp_xlsx)
        wb2.calculation.fullCalcOnLoad = True
        upsert_nav_history_row(wb2['淨值歷史'], today_iso, nav, fx, fund_mv, total_assets, total_debt, net_worth,
                                f'MoneyDJ／{fx_source}（每日自動）')
        repair_layout_and_charts(wb2)
        if '股票部位' in wb2.sheetnames:
            restyle_stock_positions_sheet(wb2['股票部位'])
        autofit_other_sheets(wb2)
        wb2.save(tmp_xlsx)

        # ── 6) 重新加密存回常駐副本 ──
        ed.save_encrypted_workbook_copy(tmp_xlsx, ENC_PATH, password)

        # ── 7) 重新產生 data.json（含健檢比對）給手機版網頁 ──
        old_data = None
        if os.path.exists(JSON_PATH):
            try:
                with open(JSON_PATH, encoding='utf-8') as f:
                    old_raw = json.load(f)
                if old_raw.get('encrypted'):
                    old_data = ed.decrypt_json(old_raw['data'], password)
            except Exception as e:
                print(f'讀取舊版 {JSON_PATH} 失敗（略過健檢比對）: {e}')

        final_data = ed.extract(tmp_xlsx)
        warnings = ed.run_health_check(final_data, old_data)
        final_data['_healthCheck'] = {
            'ok': len(warnings) == 0,
            'warnings': warnings,
            'checkedAt': final_data['generatedAt'],
        }
        ciphertext = ed.encrypt_json(final_data, password)
        with open(JSON_PATH, 'w', encoding='utf-8') as f:
            json.dump({'generatedAt': final_data['generatedAt'], 'encrypted': True, 'data': ciphertext}, f,
                      ensure_ascii=False, indent=2)
        print(f'🔒 已更新 {JSON_PATH}')
        if warnings:
            print('\n⚠️  健檢發現以下可疑異常：')
            for w in warnings:
                print(f'  - {w}')
        else:
            print('✅ 健檢通過，未發現異常')

    return 0


if __name__ == '__main__':
    sys.exit(main())
