#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自動抓取摩根多重收益基金（美元對沖）A股 當日淨值
來源依序嘗試：基富通 → 鉅亨買基金 → 保留舊值
"""
import requests, json, re, sys
from datetime import datetime, timezone, timedelta

DATA_FILE  = 'data.json'
# 基金識別
FUND_ISIN  = 'LU2347655073'
FUND_ANUE  = 'B08291'        # 鉅亨買基金代碼

HDR = {
    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) '
                  'AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}

# ── 台灣今日日期 ──────────────────────────────────────────────
TW_TZ  = timezone(timedelta(hours=8))
TODAY  = datetime.now(TW_TZ).strftime('%Y-%m-%d')

# ── 方法1：基富通 ────────────────────────────────────────────
def from_fundrich():
    try:
        url = (f'https://www.fundrich.com.tw/fund/FundDetail/'
               f'getFundDetailInfo.do?isinCode={FUND_ISIN}')
        r = requests.get(url, headers=HDR, timeout=12)
        if r.ok:
            d = r.json()
            # 嘗試常見欄位名稱
            for key in ['nav','NAV','netAssetValue','navValue','price']:
                if key in d:
                    v = float(d[key])
                    if 40 < v < 200: return v, '基富通'
            # 若是巢狀結構
            txt = r.text
            m = re.search(r'"(?:nav|NAV|price)"[:\s]+"?(\d{2,3}\.\d{1,4})"?', txt)
            if m:
                v = float(m.group(1))
                if 40 < v < 200: return v, '基富通'
    except Exception as e:
        print(f'  基富通失敗：{e}')
    return None, None

# ── 方法2：鉅亨買基金 API ─────────────────────────────────────
def from_anue_api():
    try:
        url = f'https://fund.api.cnyes.com/fund/api/v2/funds/{FUND_ANUE}/nav'
        r = requests.get(url, headers={**HDR,'Referer':'https://www.anuefund.com/'}, timeout=12)
        if r.ok:
            d = r.json()
            # 遍歷可能的資料結構
            for path in [['data','nav'],['data',0,'nav'],['nav'],['items',0,'nav']]:
                try:
                    v = d
                    for k in path: v = v[k]
                    v = float(v)
                    if 40 < v < 200: return v, '鉅亨API'
                except: pass
    except Exception as e:
        print(f'  鉅亨API失敗：{e}')
    return None, None

# ── 方法3：鉅亨網頁解析 ──────────────────────────────────────
def from_anue_page():
    try:
        url = f'https://www.anuefund.com/fund/detail/{FUND_ANUE}'
        r = requests.get(url, headers=HDR, timeout=15)
        if r.ok:
            patterns = [
                r'"nav"\s*:\s*"?(\d{2,3}\.\d{1,4})"?',
                r'最新淨值[^\d]{0,10}(\d{2,3}\.\d{1,4})',
                r'基準價格[^\d]{0,10}(\d{2,3}\.\d{1,4})',
                r'NAV[^\d]{0,10}(\d{2,3}\.\d{1,4})',
            ]
            for p in patterns:
                m = re.search(p, r.text)
                if m:
                    v = float(m.group(1))
                    if 40 < v < 200: return v, '鉅亨頁面'
    except Exception as e:
        print(f'  鉅亨頁面失敗：{e}')
    return None, None

# ── 更新 data.json ────────────────────────────────────────────
def update(nav, source):
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        data = {}

    old = data.get('nav', '無')
    data['nav']            = round(nav, 4)
    data['navSourceDate']  = TODAY
    data['navSource']      = source
    data['navFetchedAt']   = datetime.now(TW_TZ).isoformat()

    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'✅ 淨值已更新：{old} → {nav}（來源：{source}，日期：{TODAY}）')

# ── 主程式 ───────────────────────────────────────────────────
def main():
    print(f'[{TODAY}] 開始抓取基金淨值...')
    for fn in [from_fundrich, from_anue_api, from_anue_page]:
        nav, src = fn()
        if nav:
            update(nav, src)
            return
    # 全部來源失敗：不中斷流程，保留 Excel 原有淨值，正常結束（exit 0）
    print('⚠️ 所有外部來源皆抓取失敗，保留 Excel 原有淨值，不中斷同步流程')

if __name__ == '__main__':
    main()
