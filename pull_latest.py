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

會在目前資料夾產生（或覆蓋）：
    財務管理系統.xlsx      ← 最新版Excel
    pull_latest_log.txt   ← 這次執行的完整記錄（不管成功或失敗都會寫），執行完會自動用記事本打開

【為什麼要自己寫log檔、不能只看終端機視窗】
Windows的PowerShell在用 > 或 *> 把程式輸出重新導向存成檔案時，會先用主控台目前的編碼
（常常是Big5/cp950）把文字解讀一次，再用UTF-16重新編碼存檔——中文字和emoji符號經過這兩次
編碼轉換常常會變成亂碼或方塊（跟Python本身有沒有處理UTF-8無關，是PowerShell轉存那一步的問題）。
另外用滑鼠雙擊.py檔案執行時，程式一結束主控台視窗就會自動關掉，根本來不及看內容。
這裡的解法是：程式自己用UTF-8把記錄直接寫進 pull_latest_log.txt（不透過主控台、不透過
PowerShell的重新導向），寫完後自動開記事本顯示這個檔案——不管是雙擊執行還是在終端機執行，
不管主控台視窗編碼設定是什麼、會不會秒關，都能正常看到正確的中文內容。
"""
import base64
import datetime
import getpass
import hashlib
import json
import os
import subprocess
import sys
import traceback
import urllib.request

RAW_URL = 'https://raw.githubusercontent.com/eminemollie/finance/main/financial-workbook.enc'
_ENC_SALT = b'finance-dashboard-salt-v1'
OUT_FILE = '財務管理系統.xlsx'
LOG_FILE = 'pull_latest_log.txt'

_log_lines = []


def log(msg=''):
    """記一行到log（最後會整個寫進UTF-8檔案），同時盡量印到主控台方便即時查看
    （主控台印失敗——例如編碼不支援——不影響log檔本身，直接忽略即可）。"""
    _log_lines.append(msg)
    try:
        print(msg)
    except Exception:
        pass


def _write_log_and_open():
    log_path = os.path.abspath(LOG_FILE)
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(_log_lines) + '\n')
    if sys.platform == 'win32':
        try:
            subprocess.Popen(['notepad.exe', log_path])
        except Exception:
            pass


def derive_key(password: str) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), _ENC_SALT, 100000, dklen=32)


def decrypt_bytes(ciphertext_b64: str, password: str) -> bytes:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    key = derive_key(password)
    combined = base64.b64decode(ciphertext_b64)
    iv, ciphertext = combined[:12], combined[12:]
    aesgcm = AESGCM(key)
    return aesgcm.decrypt(iv, ciphertext, None)


def run():
    log(f'=== pull_latest.py 執行記錄（{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}）===')

    password = sys.argv[1] if len(sys.argv) > 1 else os.environ.get('DASHBOARD_PASSWORD')
    if not password:
        password = getpass.getpass('請輸入 DASHBOARD_PASSWORD（不會顯示在畫面上，輸入完按Enter）：')

    log(f'下載 {RAW_URL} ...')
    with urllib.request.urlopen(RAW_URL, timeout=20) as resp:
        raw = resp.read()

    payload = json.loads(raw)
    if not payload.get('encrypted'):
        log('[錯誤] 這個檔案看起來不是加密格式，請確認repo內容是否正常')
        return 1

    try:
        xlsx_bytes = decrypt_bytes(payload['data'], password)
    except Exception as e:
        log(f'[錯誤] 解密失敗，密碼可能不對：{e}')
        return 1

    with open(OUT_FILE, 'wb') as f:
        f.write(xlsx_bytes)

    log(f'[成功] 已取得最新版本並存成：{OUT_FILE}')
    log(f'       （雲端最後更新時間：{payload.get("generatedAt", "未知")}）')
    log('       接下來可以直接在這份檔案上編輯，編輯完 push 回 repo 會再觸發一次同步。')
    return 0


def main():
    try:
        exit_code = run()
    except Exception:
        log('[錯誤] 發生未預期的例外狀況：')
        log(traceback.format_exc())
        exit_code = 1
    finally:
        _write_log_and_open()
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
