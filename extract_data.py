#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_data.py
從「財務管理系統.xlsx」擷取關鍵資料，輸出 data.json 供手機版網頁讀取。
用途：GitHub Actions 在偵測到 xlsx 更新時自動執行本腳本。

執行方式：python3 extract_data.py <xlsx路徑> <輸出json路徑>
"""
import sys
import os
import json
import base64
import hashlib
import datetime
import openpyxl


# ════════════════════════════════════════════════════════════
# 端對端加密：跟網頁端（index_v3.html 的 deriveKey/decryptFromCloud）
# 使用完全相同的演算法參數，已用真實瀏覽器 Web Crypto API 驗證過
# 兩邊可以互相解密。密碼只透過 GitHub Secrets 傳入，不會寫進程式碼。
# ════════════════════════════════════════════════════════════
_ENC_SALT = b'finance-dashboard-salt-v1'


def derive_key(password: str) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), _ENC_SALT, 100000, dklen=32)


def encrypt_json(obj, password: str) -> str:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    key = derive_key(password)
    aesgcm = AESGCM(key)
    iv = os.urandom(12)
    plaintext = json.dumps(obj, ensure_ascii=False).encode('utf-8')
    ciphertext = aesgcm.encrypt(iv, plaintext, None)
    combined = iv + ciphertext
    return base64.b64encode(combined).decode('ascii')


def decrypt_json(ciphertext_b64: str, password: str):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    key = derive_key(password)
    combined = base64.b64decode(ciphertext_b64)
    iv, ciphertext = combined[:12], combined[12:]
    aesgcm = AESGCM(key)
    plaintext = aesgcm.decrypt(iv, ciphertext, None)
    return json.loads(plaintext.decode('utf-8'))


def col_letter(idx):
    letter = ''
    n = idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letter = chr(65 + rem) + letter
    return letter


def to_iso_date(v):
    """Excel 的日期值轉成 YYYY-MM-DD 字串（JS Date 可直接解析）"""
    if v is None or v == '':
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime('%Y-%m-%d')
    return None


def get_credit_card_total(wb):
    """直接加總信用卡年支出分頁「信用卡明細」區塊的原始消費金額，
    完全不依賴任何公式快取。掃描範圍僅限信用卡明細（固定+非固定）兩個表格，
    遇到小計/總計列或「每年龐大固定支出」區塊即停止，避免把小計/總計跟個別項目重複加總。"""
    try:
        ws = wb['信用卡年支出']
    except KeyError:
        return None
    total = 0
    in_cc_section = False
    _stop_markers = ('每年龐大固定支出', '龐大固定支出', '各卡消費金額合計', '信用卡費總計')
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        b_val = ws.cell(row=r, column=2).value
        if isinstance(a_val, str) and '信用卡明細' in a_val:
            in_cc_section = True
            continue
        if any(isinstance(v, str) and marker in v for v in (a_val, b_val) for marker in _stop_markers):
            break  # 信用卡區塊的小計/總計列，或整個信用卡區塊結束，停止掃描
        if not in_cc_section:
            continue
        for c in range(3, 8):  # C~G 欄：玉山/聯邦/中信/國泰/台新
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                total += v
    return abs(round(total)) if total else None


def get_jiayan_remaining(wb):
    """從股票投資分頁的買賣原始明細重新加總「市場上剩餘金額」，
    不依賴任何公式快取（原本是跨分頁公式，快取遺失時會讀不到）。"""
    try:
        ws = wb['股票投資']
    except KeyError:
        return None
    buy_total, sell_total = 0, 0
    mode = None
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        if isinstance(a_val, str) and '買入記錄' in a_val:
            mode = 'buy'; continue
        if isinstance(a_val, str) and '賣出記錄' in a_val:
            mode = 'sell'; continue
        if isinstance(a_val, str) and '共投入金額' in a_val:
            mode = None; continue
        if isinstance(a_val, str) and '共收入金額' in a_val:
            mode = None; continue
        if isinstance(a_val, str) and '小結' in a_val:
            break
        if mode in ('buy', 'sell'):
            c_val = ws.cell(row=r, column=3).value  # 台幣收入／台幣轉匯
            if isinstance(c_val, (int, float)):
                if mode == 'buy':
                    buy_total += c_val
                else:
                    sell_total += c_val
    return round(buy_total - sell_total)


def get_recent_fund_dividend_avg(wb):
    """從基金配息紀錄分頁的原始逐月實收台幣，直接在Python端重新計算最近3個月平均，
    不依賴Excel公式快取（跟收支明細裡的AVERAGE(OFFSET(...))公式邏輯一致）。"""
    if '基金配息紀錄' not in wb.sheetnames:
        return None
    ws = wb['基金配息紀錄']
    monthly_twd = []
    mode = None
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        b_val = ws.cell(row=r, column=2).value
        if a_val == '年度' and b_val == '月份':
            mode = 'hist'
            continue
        if mode == 'hist':
            f_val = ws.cell(row=r, column=6).value  # F: 實收台幣
            if isinstance(f_val, (int, float)):
                monthly_twd.append(f_val)
            if a_val == '統計摘要':
                mode = None
    if not monthly_twd:
        return None
    recent = monthly_twd[-3:] if len(monthly_twd) >= 3 else monthly_twd
    return round(sum(recent) / len(recent))


def get_childcare_avg(wb):
    """從育兒費分頁的「明細清單」（K:N欄）原始逐筆資料，直接在Python端依年月分組，
    重新計算近12個月平均本人負擔，不依賴Excel公式快取。教育費+醫療費+保險費都算，
    保費實際由每月現金流支付（不是從年終獎金另外撥款），併入計算才能真實反映月支出。
    不採用「月彙總」欄位，因為該欄位對未來尚未發生的月份也會用SUMIFS算出0（不是空白），
    直接抓會把零值月份混入平均。"""
    if '育兒費' not in wb.sheetnames:
        return None
    ws = wb['育兒費']
    month_totals = {}
    for r in range(1, ws.max_row + 1):
        month_val = ws.cell(row=r, column=11).value  # K: 年月
        cat_val = ws.cell(row=r, column=12).value     # L: 類別
        amt_val = ws.cell(row=r, column=13).value     # M: 金額
        if not isinstance(month_val, str) or cat_val not in ('教育費', '醫療費', '保險費'):
            continue
        if not isinstance(amt_val, (int, float)):
            continue
        month_totals[month_val] = month_totals.get(month_val, 0) + amt_val
    if not month_totals:
        return None
    # "113年01月"這種固定寬度格式，字串排序等同時間順序
    sorted_months = sorted(month_totals.keys())
    recent = sorted_months[-12:] if len(sorted_months) >= 12 else sorted_months
    total = sum(month_totals[m] for m in recent)
    return round(total / len(recent) / 2)


def extract(xlsx_path):
    # data_only=True 讀取快取值（用於數字），但日期/公式結構仍需從貸款表原始儲存格取得
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    result = {
        'generatedAt': datetime.datetime.now().isoformat(),
        'sourceFile': 'financial_management_system.xlsx',
    }

    # ── 收支明細：假設值 + 收入/支出項目 ─────────────────────────
    ws = wb['收支明細']
    # 配息率假設 (D欄第4列附近，尋找有百分比格式的黃底儲存格)
    rate = None
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and isinstance(c.value, (int, float)) and 0 < c.value < 1:
                if c.fill and c.fill.fgColor and c.fill.fgColor.rgb == 'FFFFFF00':
                    rate = c.value
                    break
        if rate:
            break
    if rate is None:
        rate = 0.11

    income = []
    expense = []
    section = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        d = ws.cell(row=r, column=4).value
        if isinstance(a, str) and a.startswith('💰') and '收入' in a:
            section = 'income'
            continue
        if isinstance(a, str) and a.startswith('💸') and '支出' in a:
            section = 'expense'
            continue
        if b in ('項目', None, ''):
            continue
        if b == '信用卡費' and section == 'expense':
            # 信用卡費是跨分頁公式，快取可能為空；一律改用信用卡年支出分頁的原始明細重新加總
            cc_total = get_credit_card_total(wb)
            amt = cc_total if cc_total is not None else (round(abs(c)) if isinstance(c, (int, float)) else 0)
            expense.append({'name': b, 'amt': amt, 'sub': (d or '')[:24]})
            continue
        if b == '育兒費(本人負擔)' and section == 'expense':
            # 同樣是跨分頁公式，快取可能為空；改用育兒費分頁原始明細重新計算
            cc_avg = get_childcare_avg(wb)
            amt = cc_avg if cc_avg is not None else (round(abs(c)) if isinstance(c, (int, float)) else 0)
            expense.append({'name': b, 'amt': amt, 'sub': (d or '')[:24]})
            continue
        if not isinstance(c, (int, float)):
            if section == 'income' and str(b) == '基金配息(近3個月實際平均)':
                pass  # 基金配息收入是公式，快取可能為空，下面用獨立函式重新計算，這裡先不跳過
            else:
                continue
        if section == 'income':
            is_fund = str(b) == '基金配息(近3個月實際平均)'
            if is_fund:
                fund_avg = get_recent_fund_dividend_avg(wb)
                amt = fund_avg if fund_avg is not None else (round(c) if isinstance(c, (int, float)) else 0)
                income.append({'name': b, 'amt': amt, 'auto': False})
            else:
                income.append({'name': b, 'amt': round(c), 'auto': False})
        elif section == 'expense':
            expense.append({'name': b, 'amt': round(abs(c)), 'sub': (d or '')[:24]})

    result['assumptions_rate'] = rate
    result['income'] = income
    result['expense'] = expense

    # ── 貸款總覽：貸款與壽險結構性事實 ─────────────────────────
    ws = wb['貸款總覽']
    loans = []
    insurance = []
    region_map = {'北區': 'north', '東區': 'east', '-': None}
    # 找出貸款區塊（標題列後直到「合計」列）與壽險區塊
    mode = None
    for r in range(1, ws.max_row + 1):
        b = ws.cell(row=r, column=2).value
        if b == '貸款名稱':
            mode = 'loan_header'
            continue
        if mode == 'loan_header':
            mode = 'loan'
        if b == '合計':
            mode = None
            continue
        if b == '保單名稱':
            mode = 'ins_header'
            continue
        if mode == 'ins_header':
            mode = 'ins'
        if b == '壽險合計':
            mode = None
            continue

        if mode == 'loan' and b:
            region = ws.cell(row=r, column=3).value
            principal = ws.cell(row=r, column=5).value
            rate_v = ws.cell(row=r, column=6).value
            years = ws.cell(row=r, column=7).value
            start_raw = ws.cell(row=r, column=8).value
            if isinstance(principal, (int, float)) and isinstance(rate_v, (int, float)):
                loans.append({
                    'name': b,
                    'region': region_map.get(region, None),
                    'principal': round(principal),
                    'rate': rate_v,
                    'years': years,
                    'start': to_iso_date(start_raw),
                })
        if mode == 'ins' and b:
            region = ws.cell(row=r, column=3).value
            premium = ws.cell(row=r, column=5).value
            years = ws.cell(row=r, column=6).value
            start_raw = ws.cell(row=r, column=8).value
            if isinstance(premium, (int, float)):
                insurance.append({
                    'name': b + '剩餘保費',
                    'region': region_map.get(region, None),
                    'premium': round(premium),
                    'years': years,
                    'start': to_iso_date(start_raw),
                })

    result['loans'] = loans
    result['insurance'] = insurance

    # ── 基金配息紀錄：申購批次 + 月配息歷史 ─────────────────────
    ws = wb['基金配息紀錄']
    batches = []
    mode = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a == '批次':
            mode = 'batch'
            continue
        if mode == 'batch' and isinstance(a, int):
            nav = ws.cell(row=r, column=3).value    # C: 購買點位(NAV,USD)
            units = ws.cell(row=r, column=4).value  # D: 申購單位數
            twd = ws.cell(row=r, column=6).value    # F: 單筆申購金額(TWD)
            desc = ws.cell(row=r, column=10).value  # J: 說明
            if isinstance(units, (int, float)):
                batches.append({
                    'no': a, 'date': str(b), 'desc': desc,
                    'nav': nav, 'units': units, 'twd': round(twd) if twd else 0,
                })
        if a == '目前部位小結':
            mode = None

    total_units = sum(x['units'] for x in batches)
    total_twd = sum(x['twd'] for x in batches)

    # 建立「累計單位數 -> 累計投入台幣」對照表，供月配息紀錄比對使用
    # 這樣即使 Excel 快取遺失（H欄公式結果讀不到），也能在 Python 端重新推算 cost，不依賴快取
    cum_units, cum_twd = 0.0, 0
    tier_table = []
    for batch in batches:
        cum_units += batch['units']
        cum_twd += batch['twd']
        tier_table.append((round(cum_units, 3), cum_twd))

    def cost_for_units(units_val):
        """依持有單位數比對批次累計表，回傳對應的累計投入成本(TWD)"""
        best = None
        for tier_units, tier_cost in tier_table:
            if abs(units_val - tier_units) < 0.5:
                return tier_cost
            if units_val <= tier_units:
                best = tier_cost
                break
        return best if best is not None else tier_table[-1][1]

    # 月配息歷史（cost 一律由 Python 依單位數重新推算，不讀取 Excel 公式快取）
    hist = []
    mode = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        if a == '年度' and ws.cell(row=r, column=2).value == '月份':
            mode = 'hist'
            continue
        if mode == 'hist':
            if isinstance(c, (int, float)) and b is not None:
                d_val = ws.cell(row=r, column=4).value  # 持有單位數（原始值，非公式）
                f_val = ws.cell(row=r, column=6).value  # 實收台幣（原始值，非公式）
                if isinstance(d_val, (int, float)) and isinstance(f_val, (int, float)):
                    hist.append({
                        'yr': str(a) if a else '',
                        'mo': str(b),
                        'ud': c,
                        'twd': round(f_val),
                        'units': d_val,
                        'cost': cost_for_units(d_val),
                    })
            if a == '統計摘要':
                mode = None

    result['batches'] = batches
    result['totalUnits'] = round(total_units, 3)
    result['totalInvestedTwd'] = total_twd
    result['hist'] = hist

    # ── 資產負債表：市值假設與資產預設值 ─────────────────────
    ws = wb['資產負債表']
    nav_val, fx_val = 74.91, 31.5
    assets = {'north_re': 0, 'east_re': 0, 'stocks': 0, 'jiayan': 0, 'cash': 0}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        label = b or a  # 部分列（如家妍投資備忘列）標籤因合併儲存格落在A欄
        if label and '最新淨值' in str(label) and isinstance(c, (int, float)):
            nav_val = c
        if label and '目前匯率' in str(label) and isinstance(c, (int, float)):
            fx_val = c
        if b == '北區不動產估值' and isinstance(c, (int, float)):
            assets['north_re'] = round(c)
        if b == '東區不動產估值' and isinstance(c, (int, float)):
            assets['east_re'] = round(c)
        if b and '股票市值' in str(b) and isinstance(c, (int, float)):
            assets['stocks'] = round(c)
        if label and '家妍投資剩餘部位' in str(label):
            jy = get_jiayan_remaining(wb)
            assets['jiayan'] = jy if jy is not None else (round(c) if isinstance(c, (int, float)) else 0)
        if b == '現金／存款' and isinstance(c, (int, float)):
            assets['cash'] = round(c)

    result['nav'] = nav_val
    result['fx'] = fx_val
    result['assetsDefault'] = assets

    # ── 育兒費：逐月明細 + 近12個月平均（本人負擔，不含已列計的保費）─
    # 直接從「明細清單」(K:N欄)原始逐筆資料在Python端依年月+類別分組重新加總，
    # 不讀「月彙總」B:D欄，因為那幾欄現在是SUMIFS公式，快取可能為空
    childcare_months = []
    if '育兒費' in wb.sheetnames:
        ws = wb['育兒費']
        month_data = {}  # {年月: {'edu':.., 'ins':.., 'med':..}}
        month_order = []
        for r in range(1, ws.max_row + 1):
            month_val = ws.cell(row=r, column=11).value  # K: 年月
            cat_val = ws.cell(row=r, column=12).value     # L: 類別
            amt_val = ws.cell(row=r, column=13).value     # M: 金額
            if not isinstance(month_val, str) or cat_val not in ('教育費', '保險費', '醫療費'):
                continue
            if not isinstance(amt_val, (int, float)):
                continue
            if month_val not in month_data:
                month_data[month_val] = {'edu': 0, 'ins': 0, 'med': 0}
                month_order.append(month_val)
            key = {'教育費': 'edu', '保險費': 'ins', '醫療費': 'med'}[cat_val]
            month_data[month_val][key] += amt_val
        # 依年月字串排序（固定寬度格式，字串排序等同時間順序）
        for m in sorted(month_order):
            v = month_data[m]
            childcare_months.append({'month': m, 'edu': round(v['edu']), 'ins': round(v['ins']), 'med': round(v['med'])})

    result['childcareMonths'] = childcare_months
    result['childcareAvg'] = get_childcare_avg(wb)

    return result


# ════════════════════════════════════════════════════════════
# 健檢機制：跟「上一次的 data.json」比對，抓出異常的擷取結果
# ════════════════════════════════════════════════════════════
def run_health_check(new_data, old_data):
    """比對新舊資料，找出可疑的異常變化。回傳 warnings 清單（可能為空）。"""
    warnings = []
    if old_data is None:
        return warnings  # 沒有舊資料可比對（第一次執行），略過健檢

    def count(d, key):
        v = d.get(key)
        return len(v) if isinstance(v, list) else 0

    def total_expense(d):
        return sum(x.get('amt', 0) for x in d.get('expense', []) if isinstance(x, dict))

    def total_income_nonfund(d):
        # 「基金配息(近3個月實際平均)」用精確名稱比對排除，這筆本來就會隨新增申購、
        # 實際配息表現自然波動，不適合跟月薪這種穩定收入用同一個門檻比對
        return sum(x.get('amt', 0) for x in d.get('income', [])
                   if isinstance(x, dict) and x.get('name') != '基金配息(近3個月實際平均)')

    def fund_income(d):
        for x in d.get('income', []):
            if isinstance(x, dict) and x.get('name') == '基金配息(近3個月實際平均)':
                return x.get('amt', 0)
        return None

    categories = [
        ('income', '收入項目'), ('expense', '支出項目'),
        ('loans', '貸款'), ('insurance', '壽險'),
        ('batches', '申購批次'), ('hist', '配息紀錄'),
    ]

    # 檢查1：任何原本有資料的類別，這次變成完全空白
    for key, label in categories:
        old_n, new_n = count(old_data, key), count(new_data, key)
        if old_n > 0 and new_n == 0:
            warnings.append(f'{label}從{old_n}筆變成0筆，可能是Excel排版被改動導致擷取失敗')
        elif old_n >= 3 and new_n > 0 and new_n < old_n * 0.5:
            warnings.append(f'{label}筆數從{old_n}筆大幅減少為{new_n}筆，請確認是否為預期中的變動')

    # 檢查2：關鍵金額出現異常暴增暴減（排除筆數同步增加導致的合理變化）
    old_exp, new_exp = total_expense(old_data), total_expense(new_data)
    if old_exp > 0 and new_exp > 0:
        change = abs(new_exp - old_exp) / old_exp
        if change > 0.5 and count(old_data, 'expense') == count(new_data, 'expense'):
            warnings.append(f'支出總額在筆數不變的情況下變動超過50%（{old_exp:,.0f} → {new_exp:,.0f}），請確認')

    old_inc, new_inc = total_income_nonfund(old_data), total_income_nonfund(new_data)
    if old_inc > 0 and new_inc > 0:
        change = abs(new_inc - old_inc) / old_inc
        if change > 0.5:
            warnings.append(f'非基金收入總額（月薪/加班費等）變動超過50%（{old_inc:,.0f} → {new_inc:,.0f}），請確認')

    old_fund_inc, new_fund_inc = fund_income(old_data), fund_income(new_data)
    if old_fund_inc is not None and new_fund_inc is not None:
        if old_fund_inc > 0 and new_fund_inc == 0:
            warnings.append('基金配息(近3個月實際平均)歸零，可能是「月配息追蹤紀錄」資料擷取失敗')
        elif old_fund_inc > 0:
            change = abs(new_fund_inc - old_fund_inc) / old_fund_inc
            if change > 1.0:
                warnings.append(f'基金配息(近3個月實際平均)變動超過100%（{old_fund_inc:,.0f} → {new_fund_inc:,.0f}），此項目本來就會隨新增申購或實際配息表現波動，如果是預期中的變動可忽略')

    # 檢查3：基金總投入成本或總單位數異常歸零
    if old_data.get('totalInvestedTwd', 0) > 0 and new_data.get('totalInvestedTwd', 0) == 0:
        warnings.append('基金總投入成本變成0，申購批次資料可能擷取失敗')
    if old_data.get('totalUnits', 0) > 0 and new_data.get('totalUnits', 0) == 0:
        warnings.append('基金總持有單位數變成0，申購批次資料可能擷取失敗')

    return warnings


if __name__ == '__main__':
    xlsx_path = sys.argv[1] if len(sys.argv) > 1 else 'finance_system.xlsx'
    out_path = sys.argv[2] if len(sys.argv) > 2 else 'data.json'
    password = os.environ.get('DASHBOARD_PASSWORD')  # 從 GitHub Secrets 傳入，不寫死在程式碼裡

    # 讀取舊版 data.json（若存在）供健檢比對用；若舊檔是加密格式，先解密才能比對
    old_data = None
    if os.path.exists(out_path):
        try:
            with open(out_path, encoding='utf-8') as f:
                old_raw = json.load(f)
            if old_raw.get('encrypted'):
                if password:
                    old_data = decrypt_json(old_raw['data'], password)
                else:
                    print('舊版 data.json 是加密格式，但未設定 DASHBOARD_PASSWORD，略過健檢比對')
            else:
                old_data = old_raw  # 舊版本尚未加密時的相容處理
        except Exception as e:
            print(f'讀取舊版 {out_path} 失敗（略過健檢比對）: {e}')

    data = extract(xlsx_path)
    warnings = run_health_check(data, old_data)
    data['_healthCheck'] = {
        'ok': len(warnings) == 0,
        'warnings': warnings,
        'checkedAt': data['generatedAt'],
    }

    if password:
        ciphertext = encrypt_json(data, password)
        output = {
            'generatedAt': data['generatedAt'],  # 保留在加密外層，方便網頁快速判斷是否有新版本
            'encrypted': True,
            'data': ciphertext,
        }
        print('🔒 已使用密碼加密 data.json 內容')
    else:
        output = data
        print('⚠️  未設定 DASHBOARD_PASSWORD 環境變數，data.json 將以明碼輸出（未加密）')

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f'已輸出 {out_path}')
    print(f"收入項目: {len(data['income'])}, 支出項目: {len(data['expense'])}")
    print(f"貸款: {len(data['loans'])}, 壽險: {len(data['insurance'])}")
    print(f"申購批次: {len(data['batches'])}, 配息紀錄: {len(data['hist'])}")

    if warnings:
        print('\n⚠️  健檢發現以下可疑異常：')
        for w in warnings:
            print(f'  - {w}')
        print('\n（此警示會顯示在網頁上，但不會中斷同步流程）')
    else:
        print('\n✅ 健檢通過，未發現異常')
