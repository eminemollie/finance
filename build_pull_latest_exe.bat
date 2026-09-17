@echo off
REM build_pull_latest_exe.bat
REM 只需要在這個資料夾裡雙擊這個檔案「一次」，就會把 pull_latest.py 打包成
REM pull_latest.exe，之後不用再開終端機、不用再打指令，直接雙擊 pull_latest.exe
REM 就能下載+解密最新版的 財務管理系統.xlsx。
REM
REM 這裡用 cmd.exe 的 .bat（不是 PowerShell 的 .ps1）來寫，是刻意的：
REM PowerShell 用 > 這種重新導向指令存檔時，常常會把中文字/emoji搞成亂碼
REM （已經在 pull_latest.py 本身踩過這個坑），但 cmd.exe 的重新導向是單純
REM 位元組層級複製，不會有這個問題，所以這裡選用 .bat。
REM
REM 執行完會產生：
REM   pull_latest.exe   ← 之後就雙擊這個檔案即可，不用再碰終端機
REM   build_log.txt      ← 這次打包過程的完整記錄，執行完會自動開記事本顯示
REM   build_tmp\          ← 打包過程用的暫存資料夾，可以忽略/之後刪除都沒關係

echo ============================================
echo  正在安裝打包工具 PyInstaller ...
echo  （第一次執行可能要等1~2分鐘，請耐心等候）
echo ============================================
py -m pip install --upgrade pyinstaller > build_log.txt 2>&1

echo.
echo ============================================
echo  正在把 pull_latest.py 打包成 pull_latest.exe ...
echo ============================================
py -m PyInstaller --onefile --name pull_latest --collect-all cryptography --distpath . --workpath build_tmp --specpath build_tmp pull_latest.py >> build_log.txt 2>&1

echo.
if exist pull_latest.exe (
    echo [成功] pull_latest.exe 已經產生在這個資料夾裡了！
    echo 以後只要雙擊 pull_latest.exe 就能取得最新版 財務管理系統.xlsx，不用再開終端機。
) else (
    echo [失敗] 沒有看到 pull_latest.exe，打包過程可能出了問題。
    echo 打開的 build_log.txt 裡應該會有詳細錯誤訊息，把內容回報就可以繼續排查。
)

notepad build_log.txt
echo.
echo 按任意鍵關閉這個視窗...
pause >nul
