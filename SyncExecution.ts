/**
 * 將「數位平台部」尚未有變更結果的排程，同步到「執行進度」工作表的 tblExecution，
 * 供 Power Automate + Teams 卡片通知與回報追蹤使用。
 *
 * 流程分三段，任何 Excel 寫入都發生在所有驗證通過之後：
 *   1. 讀取與唯讀驗證（來源表頭、目標表格結構）
 *   2. 純資料規劃：parseSourceSchedule → indexExistingExecutions → planSync
 *   3. 依計畫寫入：表頭遷移、建表、新增列、清除空白列、資料契約格式、日期修復
 */

type CellValue = string | number | boolean;
type SyncResult = { added: number; unchanged: number; errors: string[] };

// ===== 設定 =====
const SOURCE_SHEET = "數位平台部";
const PROGRESS_SHEET = "執行進度";
const TABLE_NAME = "tblExecution";
const INITIAL_STATUS = "待執行";
const INITIAL_CARD_STATUS = "N";

// ===== 來源欄位契約（依標題定位，使用者插入欄位不影響讀取）=====
const SOURCE_HEADERS = {
  changeId: "聯繫單號",
  schedule: "其它應注意事項",
  changeDate: "變更日期",
  result: "變更結果",
  systemId: "系統編號",
  changeItem: "變更項目",
  changeTime: "變更時間"
};

// ===== tblExecution 資料契約：欄位順序、欄寬、數值格式、換行都只在這裡定義 =====
type ColumnSpec = { header: string; width: number; numberFormat?: string; wrap?: boolean };
const EXECUTION_COLUMNS: ColumnSpec[] = [
  { header: "執行代號", width: 210, numberFormat: "@" },
  { header: "聯繫單號", width: 140, numberFormat: "@" },
  // 日期必須是 ISO 文字，Power Automate formatDateTime 才能直接解析。
  { header: "執行日期", width: 105, numberFormat: "@" },
  { header: "系統編號", width: 190 },
  { header: "變更項目", width: 260 },
  // 時間保留 Excel 數值，以 hh:mm 顯示。
  { header: "變更時間", width: 105, numberFormat: "hh:mm" },
  { header: "注意事項", width: 180, wrap: true },
  { header: "執行結果", width: 95 },
  { header: "回報人員", width: 120 },
  { header: "回報時間", width: 150, numberFormat: "yyyy-mm-dd hh:mm:ss" },
  { header: "備註", width: 220, wrap: true },
  { header: "CardSent", width: 95 },
  { header: "下線來源文字", width: 280, wrap: true }
];
const EXECUTION_HEADERS: string[] = EXECUTION_COLUMNS.map((c) => c.header);
// 舊版表頭 → 新版表頭；只有在整份表頭驗證通過後才會實際更名。
const LEGACY_HEADER_RENAMES: { from: string; to: string }[] = [{ from: "執行內容", to: "注意事項" }];

const COL_ID = EXECUTION_HEADERS.indexOf("執行代號");
const COL_CHANGE_ID = EXECUTION_HEADERS.indexOf("聯繫單號");
const COL_DATE = EXECUTION_HEADERS.indexOf("執行日期");

/** 一筆即將寫入 tblExecution 的新紀錄（人工回報欄位固定為初始值）。 */
type NewExecution = {
  executionId: string;
  changeId: string;
  date: string;
  systemId: string;
  changeItem: string;
  changeTime: string | number;
  notes: string;
  sourceText: string;
};

function toTableRow(item: NewExecution): (string | number)[] {
  const byHeader: { [header: string]: string | number } = {
    "執行代號": item.executionId,
    "聯繫單號": item.changeId,
    "執行日期": item.date,
    "系統編號": item.systemId,
    "變更項目": item.changeItem,
    "變更時間": item.changeTime,
    "注意事項": item.notes,
    "執行結果": INITIAL_STATUS,
    "回報人員": "",
    "回報時間": "",
    "備註": "",
    "CardSent": INITIAL_CARD_STATUS,
    "下線來源文字": item.sourceText
  };
  return EXECUTION_HEADERS.map((header) => {
    if (!(header in byHeader)) throw new Error(`toTableRow 缺少欄位「${header}」的對應。`);
    return byHeader[header];
  });
}

// ===== 正規化 =====
function normalizeText(value: string): string {
  return value.trim().replace(/\s+/g, " ");
}

function identity(changeId: string, date: string): string {
  return JSON.stringify([normalizeText(changeId), date]);
}

function executionIdFor(changeId: string, date: string): string {
  return `${normalizeText(changeId)}-${date.replace(/-/g, "")}`;
}

/** 轉成 yyyy-mm-dd；無法辨識時回傳空字串。 */
function normalizeDate(raw: CellValue): string {
  if (typeof raw === "boolean") return "";
  if (typeof raw === "number") {
    if (!Number.isFinite(raw) || raw < 61) return "";
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(raw) * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
  }
  const m = String(raw).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (!m) return "";
  const y = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  const d = new Date(Date.UTC(y, month - 1, day));
  if (y < 1900 || (y === 1900 && month < 3) || d.getUTCFullYear() !== y || d.getUTCMonth() !== month - 1 || d.getUTCDate() !== day) return "";
  return `${y}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

/** 轉成 Excel 時間數值（0 ≤ n < 1）；空白回傳 ""；無法辨識回傳 null。 */
function normalizeChangeTime(raw: CellValue): number | "" | null {
  const value = String(raw).trim();
  if (!value) return "";
  const m = value.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const h = Number(m[1]), min = Number(m[2]), sec = Number(m[3] || "0");
    return h < 24 && min < 60 && sec < 60 ? (h * 3600 + min * 60 + sec) / 86400 : null;
  }
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 && n < 1 ? n : null;
}

// ===== 來源排程解析（純資料）=====
type SourceMetadata = { systemId: string; changeItem: string; changeTime: string | number };
type DesiredExecution = {
  changeId: string;
  date: string;
  notes: string[];
  sourceLines: string[];
  metadata: SourceMetadata;
  sourceRowNumbers: number[];
  /** 只在實際新增此筆時才回報的資料品質提醒，避免每次同步重複洗版。 */
  insertWarnings: string[];
};
type SourceSchedule = {
  desired: Map<string, DesiredExecution>;
  activeChangeIds: Set<string>;
  completedChangeIds: Set<string>;
  errors: string[];
};

/** texts / values 皆包含標題列（第 0 列），且必須從 A1 開始。 */
function parseSourceSchedule(texts: string[][], values: CellValue[][]): SourceSchedule {
  const headers = texts[0].map((x) => x.trim());
  const col = (name: string): number => {
    const index = headers.indexOf(name);
    if (index < 0) throw new Error(`主工作表缺少「${name}」欄位。`);
    return index;
  };
  const changeIdCol = col(SOURCE_HEADERS.changeId);
  const scheduleCol = col(SOURCE_HEADERS.schedule);
  const changeDateCol = col(SOURCE_HEADERS.changeDate);
  const resultCol = col(SOURCE_HEADERS.result);
  const systemCol = col(SOURCE_HEADERS.systemId);
  const itemCol = col(SOURCE_HEADERS.changeItem);
  const timeCol = col(SOURCE_HEADERS.changeTime);

  const errors: string[] = [];
  const desired = new Map<string, DesiredExecution>();
  const activeChangeIds = new Set<string>();
  const completedChangeIds = new Set<string>();
  // 例：2026/11/20 \ 注意事項
  const schedulePattern = /^(\d{4}\/\d{1,2}\/\d{1,2})\s*\\\s*(.*)$/;

  for (let i = 1; i < texts.length; i++) {
    const row = texts[i];
    const rowNumber = i + 1;
    const changeId = normalizeText(row[changeIdCol] || "");
    if (normalizeText(row[resultCol] || "")) {
      if (changeId) completedChangeIds.add(changeId);
      continue;
    }
    if (!changeId) {
      if (row.some((v) => v.trim())) errors.push(`主表第 ${rowNumber} 列變更結果空白但缺少聯繫單號。`);
      continue;
    }
    activeChangeIds.add(changeId);

    const rowWarnings: string[] = [];
    const rawTime = values[i][timeCol];
    let changeTime = normalizeChangeTime(rawTime);
    if (changeTime === null) {
      rowWarnings.push(`主表第 ${rowNumber} 列變更時間無法辨識：${row[timeCol] || String(rawTime)}，已留空。`);
      changeTime = "";
    }
    const metadata: SourceMetadata = { systemId: row[systemCol] || "", changeItem: row[itemCol] || "", changeTime };

    const rawNotes = (row[scheduleCol] || "").trim();
    const lines = rawNotes.split(/\r?\n/).map((x) => x.trim()).filter((x) => x.length > 0);
    const parsed: { date: string; notes: string; source: string }[] = [];
    let allMatch = lines.length > 0;
    for (const line of lines) {
      const match = line.match(schedulePattern);
      const date = match ? normalizeDate(match[1]) : "";
      if (!match || !date) { allMatch = false; break; }
      parsed.push({ date, notes: match[2].trim(), source: line });
    }
    // 只要有一行不符合「日期 \ 內容」，整格視為一般備註，改用「變更日期」欄。
    const entries = allMatch ? parsed : [{
      date: normalizeDate(values[i][changeDateCol]) || normalizeDate(row[changeDateCol] || ""),
      notes: rawNotes,
      source: rawNotes
    }];

    for (const entry of entries) {
      if (!entry.date) { errors.push(`主表第 ${rowNumber} 列缺少有效變更日期，無法使用備援規則。`); continue; }
      const key = identity(changeId, entry.date);
      let record = desired.get(key);
      if (!record) {
        record = { changeId, date: entry.date, notes: [], sourceLines: [], metadata, sourceRowNumbers: [], insertWarnings: [] };
        desired.set(key, record);
      }
      if (!record.sourceRowNumbers.includes(rowNumber)) {
        if (record.sourceRowNumbers.length > 0 && !sameMetadata(record.metadata, metadata)) {
          record.insertWarnings.push(
            `主表第 ${record.sourceRowNumbers[0]}、${rowNumber} 列同為 ${changeId}／${entry.date}，但系統編號、變更項目或變更時間不同；已採用第 ${record.sourceRowNumbers[0]} 列，請人工確認。`
          );
        }
        record.sourceRowNumbers.push(rowNumber);
        for (const warning of rowWarnings) {
          if (!record.insertWarnings.includes(warning)) record.insertWarnings.push(warning);
        }
      }
      if (entry.notes && !record.notes.includes(entry.notes)) record.notes.push(entry.notes);
      if (entry.source && !record.sourceLines.includes(entry.source)) record.sourceLines.push(entry.source);
    }
  }
  return { desired, activeChangeIds, completedChangeIds, errors };
}

function sameMetadata(a: SourceMetadata, b: SourceMetadata): boolean {
  return normalizeText(a.systemId) === normalizeText(b.systemId) &&
    normalizeText(a.changeItem) === normalizeText(b.changeItem) &&
    a.changeTime === b.changeTime;
}

// ===== 既有執行進度索引（純資料）=====
type ExistingExecution = { executionId: string; changeId: string };
type ExistingIndex = {
  byKey: Map<string, ExistingExecution[]>;
  /** 所有非空執行代號，不論該列其他欄位是否有效。 */
  ids: Set<string>;
  /** 有聯繫單號但日期無效的列：無法判斷是哪一天，暫停該聯繫單的新增以免重複。 */
  blockedChangeIds: Set<string>;
  errors: string[];
};

/** firstRowNumber：資料列第一列在工作表上的列號（1-based），僅用於錯誤訊息。 */
function indexExistingExecutions(texts: string[][], dateValues: CellValue[][], firstRowNumber: number): ExistingIndex {
  const byKey = new Map<string, ExistingExecution[]>();
  const ids = new Set<string>();
  const blockedChangeIds = new Set<string>();
  const errors: string[] = [];
  texts.forEach((row, i) => {
    if (row.every((v) => !v.trim())) return;
    const rowNumber = firstRowNumber + i;
    const id = normalizeText(row[COL_ID] || "");
    const changeId = normalizeText(row[COL_CHANGE_ID] || "");
    const date = normalizeDate(dateValues[i][0]) || normalizeDate(row[COL_DATE] || "");

    // 先登記代號，再驗證其他欄位：即使本列無效，也不能再新增同代號的紀錄。
    if (id) {
      if (ids.has(id)) errors.push(`執行進度有重複執行代號：${id}。`);
      ids.add(id);
    }
    if (!changeId) {
      errors.push(`執行進度第 ${rowNumber} 列缺少聯繫單號，請人工確認。`);
      return;
    }
    if (!date) {
      blockedChangeIds.add(changeId);
      errors.push(`執行進度第 ${rowNumber} 列日期無效（${row[COL_DATE] || "空白"}），已暫停新增聯繫單 ${changeId} 的執行紀錄，請先修正。`);
      return;
    }
    if (!id) {
      errors.push(`執行進度第 ${rowNumber} 列缺少執行代號，請人工確認；已視為 ${changeId}／${date} 的既有紀錄。`);
    } else {
      const expected = executionIdFor(changeId, date);
      if (id !== expected && !(id.startsWith(expected + "-") && /^\d{3,}$/.test(id.slice(expected.length + 1)))) {
        errors.push(`執行進度第 ${rowNumber} 列執行代號與聯繫單號／日期不符：${id}。`);
      }
    }
    const key = identity(changeId, date);
    const group = byKey.get(key) || [];
    group.push({ executionId: id || `第 ${rowNumber} 列`, changeId });
    byKey.set(key, group);
  });
  return { byKey, ids, blockedChangeIds, errors };
}

// ===== 對帳規劃（純資料）=====
type SyncPlan = { newRows: NewExecution[]; unchanged: number; errors: string[] };

function planSync(schedule: SourceSchedule, existing: ExistingIndex): SyncPlan {
  const errors: string[] = [];
  const newRows: NewExecution[] = [];
  const usedIds = new Set<string>(existing.ids);
  let unchanged = 0;

  schedule.desired.forEach((item, key) => {
    const group = existing.byKey.get(key) || [];
    if (group.length > 0) {
      // 不覆寫人工回報、注意事項、卡片狀態；已存在者僅去重。
      unchanged++;
      if (group.length > 1) errors.push(`既有執行進度同一聯繫單同一天有 ${group.length} 筆：${item.changeId}／${item.date}。已保留既有資料，不覆寫舊人工填寫欄位。`);
      return;
    }
    if (existing.blockedChangeIds.has(item.changeId)) {
      errors.push(`聯繫單 ${item.changeId} 有日期無效的既有執行進度，暫不新增 ${item.date}。`);
      return;
    }
    const executionId = executionIdFor(item.changeId, item.date);
    if (usedIds.has(executionId)) { errors.push(`執行代號已存在：${executionId}，已略過新增。`); return; }
    usedIds.add(executionId);
    item.insertWarnings.forEach((w) => errors.push(w));
    newRows.push({
      executionId,
      changeId: item.changeId,
      date: item.date,
      systemId: item.metadata.systemId,
      changeItem: item.metadata.changeItem,
      changeTime: item.metadata.changeTime,
      notes: item.notes.join("\n"),
      sourceText: item.sourceLines.join("\n")
    });
  });

  // 已有結果的聯繫單視為正常結案，不對其歷史執行紀錄發出來源移除警告。
  // 其他不再出現在來源中的既有資料也不自動刪除，避免影響卡片與回報。
  existing.byKey.forEach((group, key) => {
    if (schedule.desired.has(key)) return;
    const changeId = group[0].changeId;
    if (schedule.completedChangeIds.has(changeId) || !schedule.activeChangeIds.has(changeId)) return;
    errors.push(`既有執行日期已不在目前待執行排程：${group.map((x) => x.executionId).join("、")}，已保留資料，請人工確認。`);
  });

  return { newRows, unchanged, errors };
}

function columnOf(table: ExcelScript.Table, header: string): ExcelScript.TableColumn {
  const column = table.getColumnByName(header);
  if (!column) throw new Error(`${TABLE_NAME} 找不到「${header}」欄位。`);
  return column;
}

/** 找出可安全刪除的完全空白列（公式也為空），並合併為連續區段，由下往上排列。 */
function blankRowRuns(texts: string[][], formulas: string[][]): { index: number; count: number }[] {
  const runs: { index: number; count: number }[] = [];
  for (let i = texts.length - 1; i >= 0; i--) {
    const blank = texts[i].every((v) => !v.trim()) && formulas[i].every((f) => !String(f).trim());
    if (!blank) continue;
    const last = runs[runs.length - 1];
    if (last && last.index === i + 1) { last.index = i; last.count++; } else runs.push({ index: i, count: 1 });
  }
  return runs;
}

function main(workbook: ExcelScript.Workbook): SyncResult {
  // ---- 1. 讀取與唯讀驗證 ----
  const sourceSheet = workbook.getWorksheet(SOURCE_SHEET);
  if (!sourceSheet) throw new Error(`找不到「${SOURCE_SHEET}」工作表。`);
  const sourceRange = sourceSheet.getUsedRange();
  if (!sourceRange) return { added: 0, unchanged: 0, errors: [`${SOURCE_SHEET}工作表沒有資料。`] };
  if (sourceRange.getRowIndex() !== 0 || sourceRange.getColumnIndex() !== 0) throw new Error("來源工作表標題列必須從 A1 開始。");
  const sourceTexts = sourceRange.getTexts();
  if (sourceTexts.length < 2) return { added: 0, unchanged: 0, errors: ["來源工作表沒有可同步的資料列。"] };
  const schedule = parseSourceSchedule(sourceTexts, sourceRange.getValues());

  let progressSheet = workbook.getWorksheet(PROGRESS_SHEET);
  let targetTable = workbook.getTable(TABLE_NAME);
  let pendingRenames: { from: string; to: string }[] = [];
  if (targetTable) {
    if (targetTable.getWorksheet().getName() !== PROGRESS_SHEET) throw new Error(`表格 ${TABLE_NAME} 不在「${PROGRESS_SHEET}」工作表。`);
    const actual = targetTable.getHeaderRowRange().getTexts()[0];
    pendingRenames = LEGACY_HEADER_RENAMES.filter((r) => actual.includes(r.from) && !actual.includes(r.to));
    // 以「更名後」的表頭驗證；驗證失敗時工作簿完全未被修改。
    const migrated = actual.map((h) => {
      const rename = pendingRenames.find((r) => r.from === h);
      return rename ? rename.to : h;
    });
    if (migrated.length !== EXECUTION_HEADERS.length || migrated.some((x, i) => x !== EXECUTION_HEADERS[i])) {
      throw new Error(`${TABLE_NAME} 欄位結構不符，請先人工檢查；程式不會修改既有資料。`);
    }
    progressSheet = targetTable.getWorksheet();
  } else if (progressSheet && progressSheet.getUsedRange()) {
    throw new Error(`${PROGRESS_SHEET}工作表已有資料但找不到 ${TABLE_NAME}，為避免覆蓋已停止。`);
  }

  // ---- 2. 規劃（純資料）----
  let existing: ExistingIndex = { byKey: new Map<string, ExistingExecution[]>(), ids: new Set<string>(), blockedChangeIds: new Set<string>(), errors: [] };
  let blankRuns: { index: number; count: number }[] = [];
  let existingRowCount = 0;
  if (targetTable) {
    const body = targetTable.getRangeBetweenHeaderAndTotal();
    const texts = body.getTexts();
    existingRowCount = texts.length;
    const dateValues = body.getColumn(COL_DATE).getValues();
    existing = indexExistingExecutions(texts, dateValues, body.getRowIndex() + 1);
    // 只有出現顯示空白的列時才讀公式，避免誤刪「公式結果為空字串」的列。
    if (texts.some((r) => r.every((v) => !v.trim()))) blankRuns = blankRowRuns(texts, body.getFormulas());
  }
  const plan = planSync(schedule, existing);
  const errors: string[] = [...existing.errors, ...schedule.errors, ...plan.errors];
  const added = plan.newRows.length;

  // ---- 3. 寫入 ----
  const isNewTable = !targetTable;
  if (targetTable) {
    for (const rename of pendingRenames) columnOf(targetTable, rename.from).setName(rename.to);
  } else {
    if (!progressSheet) progressSheet = workbook.addWorksheet(PROGRESS_SHEET);
    // Excel 表格至少要有一列資料列；先建立一列空白佔位列並設定格式，
    // 讓 addRows 新增的列繼承「@」等格式，日期 ISO 文字不會被自動轉成日期序號。
    const lastColumn = String.fromCharCode(64 + EXECUTION_HEADERS.length);
    progressSheet.getRange(`A1:${lastColumn}2`).setValues([EXECUTION_HEADERS, EXECUTION_HEADERS.map(() => "")]);
    targetTable = progressSheet.addTable(`A1:${lastColumn}2`, true);
    targetTable.setName(TABLE_NAME);
    existingRowCount = 1;
    blankRuns = [{ index: 0, count: 1 }];
  }
  const table: ExcelScript.Table = targetTable;

  const applyNumberFormats = () => {
    EXECUTION_COLUMNS.forEach((spec) => {
      if (spec.numberFormat) columnOf(table, spec.header).getRangeBetweenHeaderAndTotal().setNumberFormatLocal(spec.numberFormat);
    });
  };

  if (added > 0) {
    applyNumberFormats(); // 新增前設定，讓新列繼承格式
    table.addRows(-1, plan.newRows.map((item) => toTableRow(item)));
  }

  // 刪除完全空白列（含初次建表的佔位列）；新增列附加在尾端，因此既有列索引不變。
  // 表格至少保留一列資料列（Excel 限制）：零筆排程時佔位列會留著，其 CardSent 為空，不會觸發卡片。
  let remaining = existingRowCount + added;
  blankRuns.forEach((run) => {
    const count = Math.min(run.count, remaining - 1);
    if (count <= 0) return;
    table.deleteRowsAt(run.index + run.count - count, count);
    remaining -= count;
  });

  // 資料契約格式每次校正，避免人工修改造成 Power Automate 解析異常。
  applyNumberFormats();

  // 修復既有日期序號（如 46331），強制日期欄為 ISO 文字；先在記憶體轉換，再一次 setValues。
  const dateRange = columnOf(table, "執行日期").getRangeBetweenHeaderAndTotal();
  const dateValues = dateRange.getValues();
  const dateTexts = dateRange.getTexts();
  const firstRowNumber = dateRange.getRowIndex() + 1;
  let datesNeedRewrite = false;
  const normalizedDates = dateValues.map((row, i) => {
    const raw = row[0];
    if (raw === "" || raw === null || raw === undefined) return [raw];
    const normalized = normalizeDate(raw) || normalizeDate(dateTexts[i][0] || "");
    if (!normalized) {
      errors.push(`執行進度第 ${firstRowNumber + i} 列的日期無法轉成 ISO：${String(raw)}。`);
      return [raw];
    }
    if (raw !== normalized) datesNeedRewrite = true;
    return [normalized];
  });
  if (datesNeedRewrite) dateRange.setValues(normalizedDates);

  // 一次性視覺格式只在首次建立表格時設定，避免每次 Power Automate 執行都重整報表。
  if (isNewTable) {
    table.setPredefinedTableStyle("TableStyleMedium2");
    const headerFormat = table.getHeaderRowRange().getFormat();
    headerFormat.getFill().setColor("#1F4E78");
    headerFormat.getFont().setColor("#FFFFFF");
    headerFormat.getFont().setBold(true);
    headerFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    headerFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    headerFormat.setWrapText(true);
    headerFormat.setRowHeight(28);

    const dataFormat = table.getRangeBetweenHeaderAndTotal().getFormat();
    dataFormat.getFill().setColor("#FFFFFF");
    dataFormat.getFont().setColor("#000000");
    dataFormat.getFont().setBold(false);
    dataFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    dataFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    dataFormat.setRowHeight(30);

    EXECUTION_COLUMNS.forEach((spec) => {
      const column = columnOf(table, spec.header);
      column.getRange().getFormat().setColumnWidth(spec.width);
      if (spec.wrap) column.getRangeBetweenHeaderAndTotal().getFormat().setWrapText(true);
    });

    table.getWorksheet().getFreezePanes().freezeRows(1);
  }

  return { added, unchanged: plan.unchanged, errors };
}
