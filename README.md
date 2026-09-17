# 個人財務管理系統

Excel 記帳 + 自動同步 + 每日自動更新基金淨值／匯率 + 手機網頁儀表板，全程端對端加密。

## 這是什麼

一套個人財務管理系統，包含：
- **Excel 帳本**（`財務管理系統.xlsx`）：十一個分頁，涵蓋儀表板、淨值歷史、資產負債表、貸款、基金投資、股票、信用卡年支出、育兒費分帳、請款單、月收支、年度總結
- **手機網頁儀表板**（`index.html`，部署於 GitHub Pages）：五個分頁（淨值總覽／月收支／情境試算／配息紀錄／育兒費），支援跨裝置雲端同步
- **手動同步流程**（`sync-excel.yml`）：上傳 Excel 後自動轉換、加密成 `data.json`
- **每日自動更新流程**（`daily-nav-update.yml`，新增）：每天自動抓取基金最新淨值（MoneyDJ）與 USD/TWD 匯率（優先台灣銀行牌告匯率，抓不到時自動改用歐洲央行參考匯率備援），更新常駐加密 Excel 副本與 `data.json`，不需要手動開 Excel 也能讓資產淨值每天跟著市價變動

## 架構

```
財務管理系統.xlsx（手動上傳觸發，或平常放在自己電腦）
      │
      ▼ sync-excel.yml
extract_data.py（解析Excel → 加密data.json → 加密存成常駐副本 financial-workbook.enc → 刪除明碼xlsx）
      │
      ├──────────────────────────────┐
      ▼                              ▼
data.json（AES-GCM加密）      financial-workbook.enc（AES-GCM加密，常駐於repo）
      │                              ▲
      ▼                              │ 每天 09:00 (台北時間) 自動解密→更新NAV/匯率→重新加密
index.html（GitHub Pages）    daily_nav_update.py（daily-nav-update.yml 排程執行）
  讀取data.json→輸入密碼解密         資料來源：MoneyDJ（基金淨值）／台灣銀行牌告匯率CSV，
                                     失敗時自動改用歐洲央行參考匯率備援（匯率）
      │
      ▼（可選）
Cloudflare Worker（finance-sync）── 跨裝置雲端同步（同樣加密）
```

**日常編輯 Excel 前，先用 `pull_latest.py` 把常駐副本解密下載下來**（見下方），這樣才不會用本機的舊版蓋掉每日自動更新寫進去的最新淨值。

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `財務管理系統.xlsx` | 主帳本，日常記帳都在這裡改（上傳後會被自動移除明碼版本，不會留在repo裡）|
| `extract_data.py` | 把Excel轉成加密JSON的解析腳本，也負責把Excel加密存成常駐副本；Excel分頁結構若異動，這裡的欄位對應也要跟著改 |
| `daily_nav_update.py` | 每日排程執行的腳本：抓NAV/匯率 → 解密常駐Excel → 更新儲存格與淨值歷史列 → 重新加密 → 重建data.json |
| `pull_latest.py` | 在自己電腦執行，下載＋解密目前repo裡最新的常駐Excel副本，避免拿舊檔覆蓋自動更新的結果 |
| `index.html` | 手機版儀表板，純前端（無需伺服器），可直接開啟或部署到GitHub Pages |
| `data.json` | 加密後的資料，每次同步（手動或每日自動）後更新 |
| `financial-workbook.enc` | 加密後的完整Excel常駐副本，供每日自動更新流程讀寫；跟 `data.json` 用同一組密碼、同一套AES-GCM加密 |
| `.github/workflows/sync-excel.yml` | 手動上傳Excel後的同步流程 |
| `.github/workflows/daily-nav-update.yml` | 每日自動更新NAV/匯率的排程流程 |

## 日常使用

**平常想改資料（新增基金申購批次、貸款、育兒費…）：**
1. 先執行 `python3 pull_latest.py` 取得含最新NAV/匯率的版本
2. 在 `財務管理系統.xlsx` 正常記帳、新增資料
3. 上傳覆蓋 repo 裡的同名檔案（拖曳上傳或git push皆可）
4. `sync-excel.yml` 自動執行：解析 → 加密data.json → 加密存回常駐副本 → 刪除明碼Excel
5. 打開手機網頁，輸入密碼，資料就更新了

**什麼都不用做：**
- 每天台北時間早上9點，`daily-nav-update.yml` 會自動抓最新淨值/匯率，更新常駐Excel副本與手機儀表板，淨值歷史分頁也會自動多一列。

若想確認同步或每日更新是否成功，到 repo 的「Actions」分頁查看執行紀錄（綠勾＝成功，紅叉＝失敗，點進去可以看詳細錯誤訊息）。**強烈建議第一次設定完，先手動觸發一次 `daily-nav-update.yml`（workflow_dispatch）確認NAV/匯率抓取邏輯正常**，MoneyDJ／台灣銀行的網頁結構未來若改版，抓取邏輯可能需要跟著調整。

## 安全性

- `data.json` 與 `financial-workbook.enc` 都使用 AES-GCM-256 加密，同一組密碼由 GitHub Secret（`DASHBOARD_PASSWORD`）控制，不會出現在程式碼或commit紀錄裡
- **設計取捨（跟舊版不同）**：舊版本 Excel 處理完會自動從 repo 移除、完全不常駐；這次為了讓 Excel 本身也能每日自動更新，改成把**加密後**的 Excel 常駐存放在 repo（`financial-workbook.enc`）。明碼 Excel 仍然不會留在 repo，只有加密後的版本會留著，解密需要同一組密碼——多了一份加密檔案常駐，是為了換取「Excel自動更新」這個功能，若之後想改回完全不常駐，把 daily-nav-update.yml 跟常駐副本相關程式碼移除即可
- 雲端同步（Cloudflare Worker + KV）同樣全程加密，伺服器端看不到明碼內容
- 網頁每次開啟都要求輸入密碼，不記住登入狀態

## 維護注意事項

**如果調整了Excel分頁的欄位順序、區塊標題文字，一定要同步檢查 `extract_data.py` 與 `daily_nav_update.py` 的解析邏輯**——這套系統靠比對特定文字（如「月彙總」「批次」「基金最新淨值(NAV,USD)」）跟欄位位置（如「C欄=NAV」）來解析資料，Excel結構一改，解析邏輯沒跟著改就會同步失敗或抓到錯誤數字。建議每次調整完Excel結構後，實際跑一次下面的指令搭配新版Excel測試，確認沒有壞掉再上傳。

```bash
# 本地測試手動同步腳本（需要設定DASHBOARD_PASSWORD環境變數）
DASHBOARD_PASSWORD="你的密碼" python3 extract_data.py "財務管理系統.xlsx" data.json

# 本地測試每日自動更新腳本（會實際連線MoneyDJ／台灣銀行）
DASHBOARD_PASSWORD="你的密碼" python3 daily_nav_update.py
```

**MoneyDJ網頁結構若改版**，`daily_nav_update.py` 裡的 `fetch_nav_moneydj()` 會抓不到資料或抓到不合理的數字，此時腳本會印警告並直接跳過本次更新（不會寫壞任何檔案、不會讓排程失敗），但淨值就會停止每日更新，需要回來調整抓取邏輯。

**關於匯率來源**：台灣銀行官網主頁（`rate.bot.com.tw/xrt`）有機器人驗證（Radware）保護，一般程式（無法執行JS的HTTP請求）連線會被擋下、回傳一個驗證挑戰頁面而不是真正資料，這是這套系統先前放棄「Excel自動更新」的主因之一。這次改抓台灣銀行另外提供的CSV下載端點（`rate.bot.com.tw/xrt/flcsv/0/day`），實務上不一定會經過同一層驗證；但如果哪天這個CSV端點也開始被擋（回應內容變成HTML而不是CSV），`fetch_fx_bot()` 會自動印警告並改用備援來源——Frankfurter（歐洲央行每日參考匯率，公開API，不會有機器人驗證問題）——資料來源會忠實記錄在Excel的「資料來源」欄位與淨值歷史裡，不會悄悄混用。如果連備援來源都失敗，才會整個跳過本次更新。

## 授權

僅供個人使用。
