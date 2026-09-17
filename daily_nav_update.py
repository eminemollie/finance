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

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference

import extract_data as ed

MONEYDJ_URL = 'https://www.moneydj.com/funddj/ya/yp010001.djhtm?a=jfzn3'  # JPM多重收益(美元對沖)-A股(穩定月配)
BOT_CSV_URL = 'https://rate.bot.com.tw/xrt/flcsv/0/day'  # 台灣銀行牌告匯率CSV（輕量端點，主要來源）
FRANKFURTER_URL = 'https://api.frankfurter.dev/v2/rate/usd/twd'  # 歐洲央行每日參考匯率（公開API，備援來源）
ENC_PATH = 'financial-workbook.enc'
JSON_PATH = 'data.json'

NAV_MIN, NAV_MAX = 20.0, 200.0
FX_MIN, FX_MAX = 20.0, 45.0


def fetch_nav_moneydj():
    """從 MoneyDJ 淨值表撈「淨值日期」「最新淨值」這一列。抓不到／格式跑掉就丟例外。"""
    resp = requests.get(MONEYDJ_URL, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }, timeout=20)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or 'utf-8'
    soup = BeautifulSoup(resp.text, 'html.parser')

    header_cells, data_cells = None, None
    for tr in soup.find_all('tr'):
        cells = [c.get_text(strip=True) for c in tr.find_all(['td', 'th'])]
        if '淨值日期' in cells and '最新淨值' in cells:
            header_cells = cells
            nxt = tr.find_next_sibling('tr')
            if nxt:
                data_cells = [c.get_text(strip=True) for c in nxt.find_all(['td', 'th'])]
            break

    if not header_cells or not data_cells:
        raise RuntimeError('MoneyDJ 頁面結構可能已變動，找不到「淨值日期／最新淨值」表格')

    idx_date = header_cells.index('淨值日期')
    idx_nav = header_cells.index('最新淨值')
    date_str = data_cells[idx_date].strip()
    nav_str = data_cells[idx_nav].strip()

    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', date_str)
    if not m:
        raise RuntimeError(f'淨值日期格式無法解析：{date_str!r}')
    nav_date = f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'

    try:
        nav = float(nav_str)
    except ValueError:
        raise RuntimeError(f'最新淨值無法轉成數字：{nav_str!r}')

    if not (NAV_MIN <= nav <= NAV_MAX):
        raise RuntimeError(f'抓到的淨值 {nav} 超出合理範圍({NAV_MIN}~{NAV_MAX})，可能抓錯欄位，放棄更新')

    return nav, nav_date


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


def repair_layout_and_charts(wb):
    """修正「淨值歷史」分頁最初版面留下的顯示問題，並把走勢圖改成柱狀圖：
    1) A2說明文字當初沒設定自動換行(wrap_text)，文字較長時會直接往右溢出儲存格範圍，
       在Excel裡看起來像「字體超出去」；這裡改成自動換行＋足夠的列高，並把幾個容易
       擠不下中文標題的欄位稍微加寬，避免標題文字被裁切。
    2) 「淨值歷史」與「儀表板」原本是折線圖，建立當時只框選了2列資料範圍
       （標題列 + 建立當下那1筆種子資料），之後每天 upsert_nav_history_row() 新增的資料列
       並不會自動被圖表納入——圖表的資料/類別範圍是寫死在圖表定義裡的固定儲存格參照。
       結果就是資料明明越存越多，圖表卻永遠只看得到最早那1個點，1個點畫不出線，
       看起來像壞掉、空白的圖表。且只有1~2個點時折線圖本來就不好看。
    這裡改成：每天都直接把兩個圖表整個重建成「柱狀圖」（每天一根柱子，資料點少的時候
    也清楚好看），資料/類別範圍固定放寬到第5000列（大約可容納13年份的每日資料），
    之後新增的資料列會自動被涵蓋，不需要再回來調整圖表定義。
    這個函式每天都會執行一次，刻意設計成重複執行也不會壞掉或疊加錯誤
    （每次都是「整個重新產生同樣設定的圖表」，不是疊加或累積修改）。"""
    # 用5000列當上限（大約可容納13年份的每日資料），不用擔心之後又要回來調整
    CHART_MAX_ROW = 5000
    # 折線圖建立時期留下的舊圖表沒有這個問題，但保留原本的預設錨點位置以防第一次找不到既有圖表
    DEFAULT_ANCHOR = {'淨值歷史': 'J4', '儀表板': 'A40'}

    ws_hist = wb['淨值歷史']
    a2 = ws_hist['A2']
    a2.alignment = Alignment(wrap_text=True, vertical='center')
    current_height = ws_hist.row_dimensions[2].height
    if current_height is None or current_height < 34:
        ws_hist.row_dimensions[2].height = 34
    # 「基金市值(TWD)」「總資產(TWD)」「總負債(TWD)」這幾欄標題文字較長，原本寬度偏窄
    # 容易讓標題字被裁切看起來像超出格外，這裡加寬一點留呼吸空間
    for col_letter, min_width in (('D', 18), ('E', 18), ('F', 18)):
        cur = ws_hist.column_dimensions[col_letter].width
        if cur is None or cur < min_width:
            ws_hist.column_dimensions[col_letter].width = min_width

    def _rebuild_as_bar_chart(ws, sheet_name):
        existing = ws._charts
        anchor = existing[0].anchor if existing else DEFAULT_ANCHOR[sheet_name]
        height = existing[0].height if existing else 8
        width = existing[0].width if existing else (22 if sheet_name == '淨值歷史' else 24)
        ws._charts = []  # 整批清掉舊圖表（不管原本是折線圖還是柱狀圖），下面重新產生一份

        chart = BarChart()
        chart.type = 'col'  # 直立柱狀圖
        chart.style = 2
        chart.title = '淨值走勢' if sheet_name == '淨值歷史' else None
        chart.y_axis.title = '淨值 (TWD)'
        chart.x_axis.title = '日期'
        chart.height = height
        chart.width = width
        chart.gapWidth = 40  # 柱子之間留一點間距，資料點少時比較不會看起來太粗一整塊
        data_ref = Reference(ws_hist, min_col=7, min_row=4, max_row=CHART_MAX_ROW)
        cats_ref = Reference(ws_hist, min_col=1, min_row=5, max_row=CHART_MAX_ROW)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, anchor)

    for sheet_name in ('淨值歷史', '儀表板'):
        _rebuild_as_bar_chart(wb[sheet_name], sheet_name)


def main():
    password = os.environ.get('DASHBOARD_PASSWORD')
    if not password:
        print('⚠️  未設定 DASHBOARD_PASSWORD，無法解密常駐Excel副本，結束（不視為錯誤）')
        return 0

    if not os.path.exists(ENC_PATH):
        print(f'⚠️  找不到 {ENC_PATH}（可能還沒手動 push 過一次新版Excel建立常駐副本），結束')
        return 0

    # ── 1) 抓取最新 NAV / 匯率，任何一個失敗就整個放棄，不動任何檔案 ──
    try:
        nav, nav_date = fetch_nav_moneydj()
        print(f'✅ MoneyDJ 淨值：{nav}（淨值日期 {nav_date}）')
    except Exception as e:
        print(f'⚠️  抓取基金淨值失敗，略過本次自動更新：{e}')
        return 0

    try:
        fx, fx_source = fetch_fx_bot()
        print(f'✅ USD/TWD 匯率：{fx}（來源：{fx_source}）')
    except Exception as e:
        print(f'⚠️  抓取匯率失敗（含備援來源），略過本次自動更新：{e}')
        return 0

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

        # ── 3) 更新 NAV／匯率／稽核欄位（保留所有公式，不動其他任何資料）──
        wb = openpyxl.load_workbook(tmp_xlsx)
        wb.calculation.fullCalcOnLoad = True
        try:
            update_assumption_cells(wb['資產負債表'], nav, fx, updated_at_str, fx_source)
        except Exception as e:
            print(f'⚠️  {e}')
            return 0
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
