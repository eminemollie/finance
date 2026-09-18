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
    .pull_latest_state.json ← 內部用的小記錄檔，記住「上一次pull完的檔案長什麼樣子」，
                              用來偵測「你是不是在還沒commit+push的情況下又點了一次」

【2026-09-18新增：覆蓋前的安全檢查＋自動備份】
這支工具原本會「不問青紅皂白」直接把本機的 財務管理系統.xlsx 整個蓋掉，換成雲端目前的版本。
如果你（或Cowork）才剛改好本機檔案、但還沒來得及commit+push，這樣蓋下去，剛剛的修改就會
不留痕跡地消失——這件事已經實際發生過一次（育兒費請款單下拉選單的修復）。
現在改成：
  1. 執行前，如果偵測到電腦上有git，就用git檢查這個資料夾是否有「還沒commit」或
     「已commit但還沒push到GitHub」的變更；如果找不到git，退而求其次，改成比對
     「這個檔案跟上次pull完之後記錄的樣子是否一樣」。
  2. 只要偵測到「本機可能有還沒同步到雲端的修改」，就會先印出警告、暫停下來讓你確認
     （輸入 y 才會繼續，其他任何輸入都會直接中止、不覆蓋）。
  3. 不管有沒有偵測到風險，真正覆蓋之前一律先把目前的檔案備份一份
     （檔名像 財務管理系統.xlsx.backup-20260918-153000.xlsx），確保就算誤判、誤按，
     也隨時能從備份救回來。
這樣就可以放心地「隨時點pull_latest.exe」，不用再自己記得「有沒有先push」這件事——
工具自己會替你把關；真的被攔下來的話，照著畫面提示先用GitHub Desktop commit+push，
或是叫Cowork幫你確認push完了沒，再重新執行一次即可。

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
import shutil
import subprocess
import sys
import traceback
import urllib.request

RAW_URL = 'https://raw.githubusercontent.com/eminemollie/finance/main/financial-workbook.enc'
_ENC_SALT = b'finance-dashboard-salt-v1'
OUT_FILE = '財務管理系統.xlsx'
LOG_FILE = 'pull_latest_log.txt'
STATE_FILE = '.pull_latest_state.json'

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


def _hash_file(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'rb') as f:
            return hashlib.sha256(f.read()).hexdigest()
    except Exception:
        return None


def _load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_state(sha):
    if not sha:
        return
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(
                {'sha256': sha, 'pulled_at': datetime.datetime.now().isoformat()},
                f, ensure_ascii=False
            )
    except Exception:
        pass


def _git_risk_check():
    """用git檢查目前資料夾是否有「還沒commit」或「已commit但還沒push」的變更。
    回傳 (能不能判斷, 有沒有風險)。找不到git.exe、這裡不是git repo、或執行逾時，
    一律回傳 (False, False)，表示「沒辦法用git判斷，交給下一層的雜湊比對機制」。"""
    try:
        r1 = subprocess.run(
            ['git', 'status', '--porcelain'],
            capture_output=True, text=True, timeout=10
        )
        if r1.returncode != 0:
            return False, False
        dirty = bool(r1.stdout.strip())

        r2 = subprocess.run(
            ['git', 'status', '-sb'],
            capture_output=True, text=True, timeout=10
        )
        first_line = r2.stdout.splitlines()[0] if r2.stdout.strip() else ''
        ahead = 'ahead' in first_line

        return True, (dirty or ahead)
    except Exception:
        return False, False


def _confirm_overwrite():
    try:
        answer = input('確定要繼續、蓋掉本機檔案嗎？(輸入 y 繼續，其他任何輸入都會取消) ')
    except Exception:
        # 沒有互動輸入可用（例如被某些自動化方式呼叫），保守起見一律不覆蓋
        return False
    return answer.strip().lower() == 'y'


def _backup_current_file():
    if not os.path.exists(OUT_FILE):
        return
    backup_name = f'{OUT_FILE}.backup-{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}.xlsx'
    try:
        shutil.copy2(OUT_FILE, backup_name)
        log(f'       （已先把目前的檔案備份成：{backup_name}）')
    except Exception as e:
        log(f'       （備份失敗，但不影響繼續執行：{e}）')


def run():
    log(f'=== pull_latest.py 執行記錄（{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}）===')

    # --- 覆蓋前的安全檢查：避免蓋掉還沒同步到雲端的本機修改 ---
    can_check_git, git_risky = _git_risk_check()
    if can_check_git:
        if git_risky:
            log('[警告] 偵測到這個資料夾有還沒commit、或已commit但還沒push到GitHub的變更！')
            log('       現在如果繼續，會把本機的 財務管理系統.xlsx 整個蓋掉，換成雲端目前的版本，')
            log('       剛剛的修改可能就救不回來了（下面雖然還是會先備份一份，但還是建議先確認）。')
            log('       建議：先用GitHub Desktop commit+push，確認變更已經上傳，再重新執行這個工具。')
            if not _confirm_overwrite():
                log('[已取消] 使用者選擇不要覆蓋，本機檔案維持原樣。')
                return 1
            log('[使用者確認] 繼續執行，照常覆蓋。')
        else:
            log('[安全檢查] 沒有偵測到未同步的變更，可以放心繼續。')
    else:
        last_sha = _load_state().get('sha256')
        cur_sha = _hash_file(OUT_FILE)
        if last_sha and cur_sha and last_sha != cur_sha:
            log('[警告] 找不到git可以用來檢查，改用上次記錄比對：')
            log('       偵測到本機 財務管理系統.xlsx 的內容，跟上一次執行這個工具之後記錄的樣子不一樣，')
            log('       可能是你自己編輯過、或Cowork幫你改過，但不確定有沒有commit+push。')
            log('       現在如果繼續，會把這份本機檔案整個蓋掉，換成雲端目前的版本。')
            if not _confirm_overwrite():
                log('[已取消] 使用者選擇不要覆蓋，本機檔案維持原樣。')
                return 1
            log('[使用者確認] 繼續執行，照常覆蓋。')

    # --- 不管有沒有風險，覆蓋前都先備份一份現有檔案，多一層保險 ---
    _backup_current_file()

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

    _save_state(_hash_file(OUT_FILE))

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
