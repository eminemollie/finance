@echo off
REM build_pull_latest_exe.bat
REM
REM Double-click this file ONCE in this folder. It packages pull_latest.py
REM into pull_latest.exe. After that, just double-click pull_latest.exe --
REM no terminal, no typed commands, ever again.
REM
REM NOTE for future maintainers: this file is written in plain ASCII on
REM purpose. Windows cmd.exe reads .bat files using the console's active
REM codepage (often Big5/950 on Traditional Chinese Windows), and if the
REM file itself contains UTF-8 encoded Chinese text, cmd.exe can misread
REM the multi-byte sequences as stray command separators -- this caused
REM real "not recognized as an internal or external command" errors when
REM the file had Chinese REM comments and echo text. Keeping this file
REM 100%% ASCII sidesteps that whole bug class. (The Python scripts it
REM calls handle their own UTF-8 output correctly on their own -- see
REM pull_latest.py, which writes its own UTF-8 log file directly instead
REM of relying on console/redirection encoding.)
REM
REM What this produces in this folder:
REM   pull_latest.exe   - double-click this from now on
REM   build_log.txt      - full log of this packaging run (opens automatically)
REM   build_tmp\          - temp files used while packaging, safe to ignore/delete

echo ============================================
echo  Installing PyInstaller ...
echo  (first run may take 1-2 minutes, please wait)
echo ============================================
py -m pip install --upgrade pyinstaller > build_log.txt 2>&1

echo.
echo ============================================
echo  Packaging pull_latest.py into pull_latest.exe ...
echo ============================================
py -m PyInstaller --onefile --name pull_latest --collect-all cryptography --distpath . --workpath build_tmp --specpath build_tmp pull_latest.py >> build_log.txt 2>&1

echo.
if exist pull_latest.exe (
    echo [OK] pull_latest.exe was created in this folder.
    echo From now on, just double-click pull_latest.exe to get the latest xlsx.
) else (
    echo [FAILED] pull_latest.exe was not created -- something went wrong.
    echo The build_log.txt file that just opened has the detailed error.
)

notepad build_log.txt
echo.
echo Press any key to close this window...
pause >nul
