# 個人財務管理系統

Excel 記帳 + 自動同步 + 手機網頁儀表板，全程端對端加密。

## 這是什麼

一套個人財務管理系統，包含：
- **Excel 帳本**（`財務管理系統.xlsx`）：九個分頁，涵蓋月收支、貸款、基金投資、信用卡年支出、育兒費分帳、請款單、資產負債表
- **手機網頁儀表板**（`index.html`，部署於 GitHub Pages）：五個分頁（淨值總覽／月收支／情境試算／配息紀錄／育兒費），支援跨裝置雲端同步
- **自動同步流程**（GitHub Actions）：上傳 Excel 後自動轉換、加密成 `data.json`，並移除原始 Excel 檔案

## 架構

```
財務管理系統.xlsx（上傳到repo）
      │
      ▼ GitHub Actions（.github/workflows/sync-excel.yml）
extract_data.py（解析Excel → 加密 → 輸出）
      │
      ▼
data.json（AES-GCM加密，密碼保護）── 自動刪除原始xlsx
      │
      ▼
index.html（GitHub Pages部署）── 讀取data.json → 輸入密碼解密 → 顯示儀表板
      │
      ▼（可選）
Cloudflare Worker（finance-sync）── 跨裝置雲端同步（同樣加密）
```

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `財務管理系統.xlsx` | 主帳本，日常記帳都在這裡改（上傳後會被自動移除，不會留在repo裡）|
| `extract_data.py` | 把Excel轉成加密JSON的解析腳本，Excel分頁結構若異動，這裡的欄位對應也要跟著改 |
| `index.html` | 手機版儀表板，純前端（無需伺服器），可直接開啟或部署到GitHub Pages |
| `data.json` | 加密後的資料，Excel每次同步後自動更新 |
| `.github/workflows/sync-excel.yml` | 自動同步流程定義 |

## 日常使用

1. 打開 `財務管理系統.xlsx`，正常記帳、新增資料
2. 上傳覆蓋 repo 裡的同名檔案（拖曳上傳或git push皆可）
3. GitHub Actions 會自動執行：解析 → 加密 → 輸出 `data.json` → 刪除原始Excel
4. 打開手機網頁，輸入密碼，資料就更新了

若上傳後想立刻確認同步是否成功，可以到 repo 的「Actions」分頁查看執行紀錄（綠勾＝成功，紅叉＝失敗，點進去可以看詳細錯誤訊息）。

## 安全性

- `data.json` 使用 AES-GCM-256 加密，密碼由 GitHub Secret（`DASHBOARD_PASSWORD`）控制，不會出現在程式碼或commit紀錄裡
- 原始Excel檔案處理完會自動從repo移除，降低敏感資料外洩風險
- 雲端同步（Cloudflare Worker + KV）同樣全程加密，伺服器端看不到明碼內容
- 網頁每次開啟都要求輸入密碼，不記住登入狀態

## 維護注意事項

**如果調整了Excel分頁的欄位順序、區塊標題文字，一定要同步檢查 `extract_data.py` 的解析邏輯**——這套系統靠比對特定文字（如「月彙總」「批次」）跟欄位位置（如「C欄=NAV」）來解析資料，Excel結構一改，解析邏輯沒跟著改就會同步失敗或抓到錯誤數字。建議每次調整完Excel結構後，實際跑一次 `extract_data.py` 搭配新版Excel測試，確認沒有壞掉再上傳。

```bash
# 本地測試同步腳本（需要設定DASHBOARD_PASSWORD環境變數）
DASHBOARD_PASSWORD="你的密碼" python3 extract_data.py "財務管理系統.xlsx" data.json
```

## 授權

僅供個人使用。
