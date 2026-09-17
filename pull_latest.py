#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pull_latest.py — 在自己電腦上執行，把 repo 裡最新的「常駐加密Excel副本」
(financial-workbook.enc，每天會被自動更新流程或手動push更新) 下載＋解密回本機的
財務管理系統.xlsx，讓你在編輯新的一批基金申購/貸款/育兒費資料前，先拿到含有
最新NAV/匯率的版本，避免拿舊的本機檔案覆蓋掉自動更新的結果。

用法：
    export DASHBOARD_PASSWORD='你的密碼'
    python3 pull_latest.py
    # 或直接：python3 pull_latest.py 你的密碼

會在目前資料夾產生（或覆蓋）：財務管理系統.xlsx
"""
import base64
import getpass
import hashlib
import os
import sys
import urllib.request

RAW_URL = 'https://raw.githubusercontent.com/eminemollie/finance/main/financial-workbook.enc'
_ENC_SALT = b'finance-dashboard-salt-v1'
OUT_FILE = '財務管理系統.xlsx'


def derive_key(password: str) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), _ENC_SALT, 100000, dklen=32)


def decrypt_bytes(ciphertext_b64: str, password: str) -> bytes:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    key = derive_key(password)
    combined = base64.b64decode(ciphertext_b64)
    iv, ciphertext = combined[:12], combined[12:]
    aesgcm = AESGCM(key)
    return aesgcm.decrypt(iv, ciphertext, None)


def main():
    password = sys.argv[1] if len(sys.argv) > 1 else os.environ.get('DASHBOARD_PASSWORD')
    if not password:
        password = getpass.getpass('請輸入 DASHBOARD_PASSWORD（不會顯示在畫面上）：')

    print(f'下載 {RAW_URL} ...')
    with urllib.request.urlopen(RAW_URL, timeout=20) as resp:
        raw = resp.read()

    import json
    payload = json.loads(raw)
    if not payload.get('encrypted'):
        print('⚠️  這個檔案看起來不是加密格式，請確認repo內容是否正常')
        sys.exit(1)

    try:
        xlsx_bytes = decrypt_bytes(payload['data'], password)
    except Exception as e:
        print(f'❌ 解密失敗，密碼可能不對：{e}')
        sys.exit(1)

    with open(OUT_FILE, 'wb') as f:
        f.write(xlsx_bytes)

    print(f'✅ 已取得最新版本並存成：{OUT_FILE}')
    print(f'   （雲端最後更新時間：{payload.get("generatedAt", "未知")}）')
    print('   接下來可以直接在這份檔案上編輯，編輯完 push 回 repo 會再觸發一次同步。')


if __name__ == '__main__':
    main()
