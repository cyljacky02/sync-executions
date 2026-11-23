# SyncExecution.ts — Code Review

日期：2026-11-20  
審查範圍：`SyncExecution.ts`（303 行）。

> 這是單一檔案的靜態與本機模擬審查。沒有 Git repository、固定比較點、專案規範文件或正式功能規格，因此**不是**以 `git diff <base>...HEAD` 進行的正式雙軸 PR 審查。並未連線執行 Excel Online 或 Power Automate。

## 可確認的功能風險

| 嚴重度 | 程式行號 | 現象 | 建議 |
|---|---|---|---|
| P1 | 84–100, 145–159 | 既有紀錄日期無效時，`return` 會使執行代號未寫入 `existingIds`；同 ID 的新紀錄可被再次加入。 | 先獨立保留所有非空執行代號；對日期不合法的既有紀錄採保守封鎖或回報衝突，避免新增重複 ID。 |
| P2 | 173–179, 236–245 | 初次建立表格固定填入一行全空白，當沒有新增排程時仍保留一行。 | 明確規定零筆排程時是否建表、是否允許零資料列；避免虛構可供下游讀取的紀錄。 |
| P2 | 127–141 | `desired` 以「聯繫單號＋日期」合併；同日不同系統、變更項目等 metadata 直接採第一次資料，沒有警告。 | 明確確認唯一性規則；資料衝突時保留診斷而非無聲採第一筆。 |
| P2 | 68–71 | 先將 `執行內容` 改名為 `注意事項`，然後才檢查其他表頭是否一致；出錯時已寫入部分修改。 | 用唯讀預檢比對整份表頭，只有全數確認後才更名或正式同步。 |
| P2（條件式） | 238–245 | 只檢查 `getTexts()` 的顯示文字是否為空；可能將公式回傳空字串的資料列當成空白列刪掉。 | 只移除本次建立的佔位列，或同時驗證原始值／公式，而非一般性刪除顯示空白資料列。 |
| P3 | 236–245 | 對每一空白列單獨呼叫 `deleteRowsAt()`，如果歷史資料較多將增加 Office Scripts 與 Excel 的往返。 | 改為只追蹤本次新增的佔位列，避免每次全表掃描與逐列刪除。 |

## 本機模擬確認

測試時使用 JavaScript 模擬 ExcelScript 工作簿與表格行為，執行由 TypeScript 轉譯的同一份 `main()`；**不是 Excel Online 實際運行**。

1. **原始表頭正常、來源只有已完成任務：** 回傳 `added: 0`、`errors: []`，但新 `tblExecution` 內仍有一行 13 欄均為空字串。
2. **既有列 ID = `X-20261120`、日期 = `broken date`；來源要同步 X 在 2026-11-20：** 回傳 `added: 1`，表內最終出現兩行相同執行 ID。
3. **既有表頭包含 `執行內容` 且其他欄位不一致：** `main()` 拋出欄位結構不符的錯誤，但 `執行內容` 已改名成 `注意事項`。
4. **來源有兩列同一聯繫單/日期但系統編號不同：** 回傳 `added: 1`、`errors: []`，只保存第一筆的系統與變更項目。

TypeScript 檢查：以本機最小 ExcelScript 型別宣告跑過 `tsc --strict --noEmit`，沒有語法／該宣告涵蓋範圍的型別錯誤；**不代表使用 Microsoft 正式 Office Scripts 編譯器或完成 Excel 動態驗證**。

## Standards（規範軸）

未找到專案的 `CODING_STANDARDS.md`、`CONTRIBUTING.md`、`GLOSSARY.md`、`docs/adr/`，以下僅為 **Fowler code smell 判斷性提示**，不是違反已知專案規範：

| Smell | 行號 | 觀察 |
|---|---|---|
| Divergent Change | 5–303 | `main` 同時處理來源讀取、排程解析、對帳、寫入、修復與格式，測試一項規則需準備整個活頁簿。 |
| Primitive Obsession / Data Clumps | 3, 117, 159 | `metadata: (string \| number)[]` 依欄位位置攜帶系統、變更項目與時間；單純變更索引可造成無聲錯位。 |
| Shotgun Surgery | 11, 84–99, 159, 263–294 | 表頭文字、位置索引及欄寬散布數個區塊；調整表格欄位要修改多個位置。 |

規範軸共有 **3 個判斷性提示**，沒有可確認的正式規範違規。

## Spec（規格軸）

原始來源是 6 張 Office Script 程式截圖，沒有附功能驗收標準、議題單、需求文件或 Git 比較基準。截圖可證實程式有重複 ID 判斷、保留人工回報與日期轉換的意圖，但**無法以此確認所有預期業務行為**。因此規格軸記為 **無正式 spec，無法評定規格符合度**；以上功能風險仍是可被程式行為證實的問題，不因缺少 spec 而消失。

## Architecture 方向

建議先深化「執行進度對帳 module」，其次是「工作簿驗證與寫入 module」，再檢視「來源排程解析 module」及「表格資料契約 module」。架構報告含前後對照示意圖，另附 HTML 檔。

**參考：** Microsoft Learn, [ExcelScript.Table](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts)；[Improve Office Scripts performance](https://learn.microsoft.com/en-us/office/dev/scripts/develop/web-client-performance)；[Power Automate troubleshooting](https://learn.microsoft.com/en-us/office/dev/scripts/testing/power-automate-troubleshooting)。
