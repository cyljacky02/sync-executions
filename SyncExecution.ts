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
const COL_SOURCE = EXECUTION_HEADERS.indexOf("下線來源文字");

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

/** 同一聯繫單同一天：第一筆為 CR-1-20261121，之後依序為 -002、-003… */
function nextExecutionId(changeId: string, date: string, usedIds: Set<string>): string {
  const base = executionIdFor(changeId, date);
  if (!usedIds.has(base)) return base;
  for (let n = 2; ; n++) {
    const candidate = `${base}-${String(n).padStart(3, "0")}`;
    if (!usedIds.has(candidate)) return candidate;
  }
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
/** 一筆應存在的執行追蹤：排程格式的一行，或一般注意事項的整格。 */
type DesiredExecution = {
  changeId: string;
  date: string;
  notes: string;
  /** 寫入「下線來源文字」，也是與既有列對帳時的比對依據。 */
  sourceText: string;
  metadata: SourceMetadata;
  sourceRowNumber: number;
  /** 只在實際新增此筆時才回報的資料品質提醒，避免每次同步重複洗版。 */
  insertWarnings: string[];
};
/** 供除錯用的來源解析統計。 */
type SourceStats = {
  dataRows: number;
  blankRows: number;
  completedRows: number;
  missingChangeIdRows: number;
  activeRows: number;
  scheduleLineEntries: number;
  fallbackEntries: number;
  invalidDateEntries: number;
  formatWarnings: number;
};
type SourceSchedule = {
  /** key = identity(聯繫單號, 日期)；value 依來源列、行的順序排列。 */
  desired: Map<string, DesiredExecution[]>;
  desiredCount: number;
  activeChangeIds: Set<string>;
  completedChangeIds: Set<string>;
  errors: string[];
  /** 各來源欄位的欄索引（0-based），依 SOURCE_HEADERS 的 key。 */
  columns: { [key: string]: number };
  stats: SourceStats;
};

// ===== 「其它應注意事項」排程格式 =====
// 每行「yyyy/m/d 內容」，日期與內容以空白分隔（內容可省略），例：
//   2026/11/21 A機房切換
//   2026/11/22 B機房切換
// 整格每一行都符合時，每一行各自是一筆執行追蹤（同一天的多行也分開）；
// 同一行內的多個項目（如「2026/11/11 機房A、機房B」）是同一筆。同一格中完全相同的重複行只計一次。
// 空白或不符合格式時，整格視為一般注意事項，以「變更日期」建立單筆追蹤。
// 相容舊格式「yyyy/m/d \ 內容」（半形或全形反斜線）。
// 注意：以日期開頭的單行備註（如「2026/11/21 前完成備份」）也符合格式，會以該日為執行日期。
const SCHEDULE_LINE = /^(\d{4}\/\d{1,2}\/\d{1,2})(?:\s*[\\＼]\s*|\s+|$)(.*)$/;
// 行首「看起來像日期」：用來判斷不符合格式的格子是否可能只是打錯，需要提醒。
// 不含「.」，避免 IP（如 10.1.1.1）或版本號開頭的一般備註被誤報。
const DATE_LIKE_START = /^\d{2,4}\s*[\/\-／年]\s*\d{1,2}/;

type ScheduleLine = { date: string; notes: string; source: string };

/**
 * 解析一格「其它應注意事項」。
 * lines 非空：全部符合排程格式。lines 為空：應整格視為一般注意事項。
 * problem 非空：格子看似想用排程格式但有行不符合，內容為不符合的原因。
 */
function parseScheduleCell(raw: string): { lines: ScheduleLine[]; problem: string } {
  const lines = raw.split(/\r?\n/).map((x) => x.trim()).filter((x) => x.length > 0);
  const parsed: ScheduleLine[] = [];
  const problems: string[] = [];
  let looksIntended = false;
  lines.forEach((line, k) => {
    const shown = line.length > 40 ? `${line.slice(0, 40)}…` : line;
    const match = line.match(SCHEDULE_LINE);
    if (match) {
      const date = normalizeDate(match[1]);
      looksIntended = true;
      if (date) parsed.push({ date, notes: match[2].trim(), source: line });
      else problems.push(`第 ${k + 1} 行「${shown}」的日期 ${match[1]} 不存在`);
    } else if (DATE_LIKE_START.test(line)) {
      looksIntended = true;
      problems.push(`第 ${k + 1} 行「${shown}」的日期須為西元 yyyy/m/d，且日期後要空一格再寫內容`);
    } else {
      problems.push(`第 ${k + 1} 行「${shown}」不是「日期 內容」格式`);
    }
  });
  if (lines.length > 0 && problems.length === 0) return { lines: parsed, problem: "" };
  const problem = looksIntended ? problems.slice(0, 3).join("；") + (problems.length > 3 ? ` 等 ${problems.length} 行` : "") : "";
  return { lines: [], problem };
}

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
  const columns: { [key: string]: number } = {
    changeId: changeIdCol, schedule: scheduleCol, changeDate: changeDateCol, result: resultCol,
    systemId: systemCol, changeItem: itemCol, changeTime: timeCol
  };
  const stats: SourceStats = {
    dataRows: texts.length - 1, blankRows: 0, completedRows: 0, missingChangeIdRows: 0,
    activeRows: 0, scheduleLineEntries: 0, fallbackEntries: 0, invalidDateEntries: 0, formatWarnings: 0
  };

  const errors: string[] = [];
  const desired = new Map<string, DesiredExecution[]>();
  let desiredCount = 0;
  const activeChangeIds = new Set<string>();
  const completedChangeIds = new Set<string>();

  for (let i = 1; i < texts.length; i++) {
    const row = texts[i];
    const rowNumber = i + 1;
    const changeId = normalizeText(row[changeIdCol] || "");
    if (row.every((v) => !v.trim())) { stats.blankRows++; continue; }
    if (normalizeText(row[resultCol] || "")) {
      stats.completedRows++;
      if (changeId) completedChangeIds.add(changeId);
      continue;
    }
    if (!changeId) {
      stats.missingChangeIdRows++;
      errors.push(`主表第 ${rowNumber} 列變更結果空白但缺少聯繫單號。`);
      continue;
    }
    stats.activeRows++;
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
    const cell = parseScheduleCell(rawNotes);
    const usesSchedule = cell.lines.length > 0;
    const fallbackDate = normalizeDate(values[i][changeDateCol]) || normalizeDate(row[changeDateCol] || "");
    if (cell.problem) {
      stats.formatWarnings++;
      rowWarnings.push(`主表第 ${rowNumber} 列「${SOURCE_HEADERS.schedule}」看似排程格式「yyyy/m/d 內容」，但${cell.problem}；已整格視為一般注意事項，以變更日期 ${fallbackDate || "(無效)"} 建立單筆追蹤。`);
    }
    // 空白或不符合格式：整格視為一般注意事項，改用「變更日期」欄。
    const entries: ScheduleLine[] = usesSchedule ? cell.lines : [{ date: fallbackDate, notes: rawNotes, source: rawNotes }];
    if (usesSchedule) stats.scheduleLineEntries += entries.length; else stats.fallbackEntries++;

    const seenLines = new Set<string>();
    for (const entry of entries) {
      if (!entry.date) {
        stats.invalidDateEntries++;
        errors.push(`主表第 ${rowNumber} 列缺少有效變更日期，無法使用備援規則。`);
        // 未建立紀錄時輸入提醒不會隨新增回報，直接列入訊息。
        rowWarnings.forEach((w) => errors.push(w));
        continue;
      }
      const lineKey = JSON.stringify([entry.date, normalizeText(entry.source)]);
      if (seenLines.has(lineKey)) continue;
      seenLines.add(lineKey);
      const key = identity(changeId, entry.date);
      const group = desired.get(key) || [];
      group.push({
        changeId, date: entry.date, notes: entry.notes, sourceText: entry.source,
        metadata, sourceRowNumber: rowNumber, insertWarnings: rowWarnings.slice()
      });
      desired.set(key, group);
      desiredCount++;
    }
  }
  return { desired, desiredCount, activeChangeIds, completedChangeIds, errors, columns, stats };
}

// ===== 既有執行進度索引（純資料）=====
type ExistingExecution = { executionId: string; changeId: string; sourceText: string };
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
    group.push({ executionId: id || `第 ${rowNumber} 列`, changeId, sourceText: row[COL_SOURCE] || "" });
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

  // 同一聯繫單同一天可能有多筆：先依「下線來源文字」完全相同配對，
  // 剩下的再依順序配對（來源只改了內容文字時仍視為同一筆，不會重複新增），
  // 仍沒有配對的來源才新增；仍沒有配對的既有列則提醒人工確認。
  // 已存在者不覆寫人工回報、注意事項、卡片狀態。
  const leftovers: ExistingExecution[] = [];
  schedule.desired.forEach((items, key) => {
    const remaining = (existing.byKey.get(key) || []).slice();
    const unmatched: DesiredExecution[] = [];
    for (const item of items) {
      const index = remaining.findIndex((e) => normalizeText(e.sourceText) === normalizeText(item.sourceText));
      if (index >= 0) { remaining.splice(index, 1); unchanged++; } else unmatched.push(item);
    }
    for (const item of unmatched) {
      if (remaining.length > 0) { remaining.shift(); unchanged++; continue; }
      if (existing.blockedChangeIds.has(item.changeId)) {
        errors.push(`聯繫單 ${item.changeId} 有日期無效的既有執行進度，暫不新增 ${item.date}。`);
        continue;
      }
      const executionId = nextExecutionId(item.changeId, item.date, usedIds);
      usedIds.add(executionId);
      // 同一來源列拆出的多筆會帶相同提醒，只列一次。
      item.insertWarnings.forEach((w) => { if (!errors.includes(w)) errors.push(w); });
      newRows.push({
        executionId,
        changeId: item.changeId,
        date: item.date,
        systemId: item.metadata.systemId,
        changeItem: item.metadata.changeItem,
        changeTime: item.metadata.changeTime,
        notes: item.notes,
        sourceText: item.sourceText
      });
    }
    remaining.forEach((e) => leftovers.push(e));
  });
  existing.byKey.forEach((group, key) => {
    if (!schedule.desired.has(key)) group.forEach((e) => leftovers.push(e));
  });

  // 已有結果的聯繫單視為正常結案，不對其歷史執行紀錄發出來源移除警告。
  // 其他不再出現在來源中的既有資料也不自動刪除，避免影響卡片與回報。
  const stale = leftovers.filter((e) => schedule.activeChangeIds.has(e.changeId) && !schedule.completedChangeIds.has(e.changeId));
  if (stale.length > 0) {
    errors.push(`既有執行紀錄已不在目前待執行排程：${stale.map((x) => x.executionId).join("、")}，已保留資料，請人工確認。`);
  }

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

// ===== 除錯紀錄 =====
// DEBUG = true：在 Office Scripts「輸出」窗格逐步印出進度與統計。
// 每次 console.log 都會與活頁簿同步一次，因此出錯時能準確對應到最後印出的步驟；
// 穩定後可改為 false，只保留最後的結果摘要。
const DEBUG = true;
const LOG_SAMPLE_LIMIT = 5;
let runStartedAt = 0;
let currentStep = "啟動";

function log(message: string): void {
  console.log(`[SyncExecution +${Date.now() - runStartedAt}ms] ${message}`);
}

function debug(message: string): void {
  if (DEBUG) log(message);
}

function step(name: string): void {
  currentStep = name;
  debug(`>> ${name}`);
}

function columnLetter(index: number): string {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function sample(items: string[]): string {
  if (items.length === 0) return "(無)";
  const shown = items.slice(0, LOG_SAMPLE_LIMIT).join("、");
  return items.length > LOG_SAMPLE_LIMIT ? `${shown} …等 ${items.length} 筆` : shown;
}

function logResult(result: SyncResult): void {
  log(`完成：新增 ${result.added} 筆、既有未變 ${result.unchanged} 筆、訊息 ${result.errors.length} 則。`);
  result.errors.forEach((message, i) => log(`  訊息 ${i + 1}/${result.errors.length}：${message}`));
}

function main(workbook: ExcelScript.Workbook): SyncResult {
  runStartedAt = Date.now();
  currentStep = "啟動";
  try {
    const result = runSync(workbook);
    logResult(result);
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    log(`[失敗] 步驟「${currentStep}」：${message}`);
    throw error;
  }
}

function runSync(workbook: ExcelScript.Workbook): SyncResult {
  // ---- 1. 讀取與唯讀驗證 ----
  step(`讀取來源工作表「${SOURCE_SHEET}」`);
  const sourceSheet = workbook.getWorksheet(SOURCE_SHEET);
  if (!sourceSheet) {
    const names = workbook.getWorksheets().map((s) => s.getName());
    throw new Error(`找不到「${SOURCE_SHEET}」工作表。目前工作表：${names.join("、")}`);
  }
  // valuesOnly = true：只計算有值的儲存格，避免整欄格式化造成數十萬列的讀取。
  const sourceRange = sourceSheet.getUsedRange(true);
  if (!sourceRange) return { added: 0, unchanged: 0, errors: [`${SOURCE_SHEET}工作表沒有資料。`] };
  debug(`來源範圍 ${sourceRange.getAddress()}（${sourceRange.getRowCount()} 列 × ${sourceRange.getColumnCount()} 欄）`);
  if (sourceRange.getRowIndex() !== 0 || sourceRange.getColumnIndex() !== 0) {
    throw new Error(`來源工作表標題列必須從 A1 開始，目前有值的範圍是 ${sourceRange.getAddress()}。`);
  }
  const sourceTexts = sourceRange.getTexts();
  const sourceValues = sourceRange.getValues();
  if (sourceTexts.length < 2) return { added: 0, unchanged: 0, errors: ["來源工作表沒有可同步的資料列。"] };
  debug(`來源標題：${sourceTexts[0].map((h, i) => `${columnLetter(i)}=${h.trim() || "(空白)"}`).join("，")}`);

  step("解析來源排程");
  const schedule = parseSourceSchedule(sourceTexts, sourceValues);
  const s = schedule.stats;
  debug(`欄位對應：${Object.keys(schedule.columns).map((k) => `${SOURCE_HEADERS[k as keyof typeof SOURCE_HEADERS]}=${columnLetter(schedule.columns[k])}`).join("，")}`);
  debug(`來源 ${s.dataRows} 列：空白 ${s.blankRows}、已有變更結果 ${s.completedRows}、缺聯繫單號 ${s.missingChangeIdRows}、待執行 ${s.activeRows}`);
  debug(`待執行列解析：「日期 內容」排程 ${s.scheduleLineEntries} 筆、改用變更日期 ${s.fallbackEntries} 筆（其中疑似格式錯誤 ${s.formatWarnings} 筆）、日期無效 ${s.invalidDateEntries} 筆 → 產生 ${schedule.desiredCount} 筆執行意圖`);
  const desiredKeys: string[] = [];
  schedule.desired.forEach((items) => items.forEach((item) => desiredKeys.push(`${item.changeId}／${item.date}／${item.notes || "(無內容)"}`)));
  debug(`執行意圖範例：${sample(desiredKeys)}`);
  if (schedule.desiredCount === 0) logEmptyScheduleHints(sourceTexts, sourceValues, schedule);

  step(`檢查目標表格 ${TABLE_NAME}`);
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
      throw new Error(`${TABLE_NAME} 欄位結構不符，請先人工檢查；程式不會修改既有資料。目前表頭：${actual.join("、")}`);
    }
    progressSheet = targetTable.getWorksheet();
    debug(`找到既有表格，範圍 ${targetTable.getRange().getAddress()}；待更名表頭：${sample(pendingRenames.map((r) => `${r.from}→${r.to}`))}`);
  } else if (progressSheet && progressSheet.getUsedRange()) {
    throw new Error(`${PROGRESS_SHEET}工作表已有資料但找不到 ${TABLE_NAME}，為避免覆蓋已停止。`);
  } else {
    debug(progressSheet ? `「${PROGRESS_SHEET}」工作表是空的，將建立新表格。` : `沒有「${PROGRESS_SHEET}」工作表，將新增工作表與表格。`);
  }

  // ---- 2. 規劃（純資料）----
  step("讀取既有執行進度");
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
  debug(`既有 ${existingRowCount} 列：執行代號 ${existing.ids.size} 個、聯繫單／日期組合 ${existing.byKey.size} 個、暫停新增的聯繫單 ${existing.blockedChangeIds.size} 個、可刪除空白區段 ${blankRuns.length} 個`);

  step("規劃同步");
  const plan = planSync(schedule, existing);
  const errors: string[] = [...existing.errors, ...schedule.errors, ...plan.errors];
  const added = plan.newRows.length;
  debug(`計畫：新增 ${added} 筆、既有未變 ${plan.unchanged} 筆、訊息 ${errors.length} 則；新增代號：${sample(plan.newRows.map((r) => r.executionId))}`);

  // ---- 3. 寫入 ----
  const isNewTable = !targetTable;
  if (targetTable) {
    if (pendingRenames.length > 0) {
      step("寫入：遷移舊表頭");
      for (const rename of pendingRenames) columnOf(targetTable, rename.from).setName(rename.to);
    }
  } else {
    step("寫入：建立工作表與表格");
    if (!progressSheet) progressSheet = workbook.addWorksheet(PROGRESS_SHEET);
    // Excel 表格至少要有一列資料列；先建立一列空白佔位列並設定格式，
    // 讓 addRows 新增的列繼承「@」等格式，日期 ISO 文字不會被自動轉成日期序號。
    const tableAddress = `A1:${columnLetter(EXECUTION_HEADERS.length - 1)}2`;
    progressSheet.getRange(tableAddress).setValues([EXECUTION_HEADERS, EXECUTION_HEADERS.map(() => "")]);
    targetTable = progressSheet.addTable(tableAddress, true);
    targetTable.setName(TABLE_NAME);
    existingRowCount = 1;
    blankRuns = [{ index: 0, count: 1 }];
    debug(`已建立 ${TABLE_NAME}（${tableAddress}，含一列佔位列）`);
  }
  const table: ExcelScript.Table = targetTable;

  const applyNumberFormats = () => {
    EXECUTION_COLUMNS.forEach((spec) => {
      if (spec.numberFormat) columnOf(table, spec.header).getRangeBetweenHeaderAndTotal().setNumberFormatLocal(spec.numberFormat);
    });
  };

  if (added > 0) {
    step(`寫入：新增 ${added} 列`);
    applyNumberFormats(); // 新增前設定，讓新列繼承格式
    table.addRows(-1, plan.newRows.map((item) => toTableRow(item)));
    debug(`新增完成，表格現有 ${table.getRowCount()} 列`);
  } else {
    debug("沒有需要新增的列。");
  }

  // 刪除完全空白列（含初次建表的佔位列）；新增列附加在尾端，因此既有列索引不變。
  // 表格至少保留一列資料列（Excel 限制）：零筆排程時佔位列會留著，其 CardSent 為空，不會觸發卡片。
  if (blankRuns.length > 0) {
    step("寫入：刪除空白列");
    let remaining = existingRowCount + added;
    let deleted = 0;
    blankRuns.forEach((run) => {
      const count = Math.min(run.count, remaining - 1);
      if (count <= 0) return;
      table.deleteRowsAt(run.index + run.count - count, count);
      remaining -= count;
      deleted += count;
    });
    debug(deleted > 0 ? `已刪除 ${deleted} 列空白列` : "保留唯一的空白佔位列（Excel 表格至少需一列資料列）");
  }

  // 資料契約格式每次校正，避免人工修改造成 Power Automate 解析異常。
  step("寫入：校正欄位數值格式");
  applyNumberFormats();

  // 修復既有日期序號（如 46331），強制日期欄為 ISO 文字；先在記憶體轉換，再一次 setValues。
  step("寫入：檢查並修復執行日期");
  const dateRange = columnOf(table, "執行日期").getRangeBetweenHeaderAndTotal();
  const dateValues = dateRange.getValues();
  const dateTexts = dateRange.getTexts();
  const firstRowNumber = dateRange.getRowIndex() + 1;
  let repairedDates = 0;
  const normalizedDates = dateValues.map((row, i) => {
    const raw = row[0];
    if (raw === "" || raw === null || raw === undefined) return [raw];
    const normalized = normalizeDate(raw) || normalizeDate(dateTexts[i][0] || "");
    if (!normalized) {
      errors.push(`執行進度第 ${firstRowNumber + i} 列的日期無法轉成 ISO：${String(raw)}。`);
      return [raw];
    }
    if (raw !== normalized) repairedDates++;
    return [normalized];
  });
  if (repairedDates > 0) dateRange.setValues(normalizedDates);
  debug(`日期修復 ${repairedDates} 格`);

  // 一次性視覺格式只在首次建立表格時設定，避免每次 Power Automate 執行都重整報表。
  if (isNewTable) {
    step("寫入：首次建表版面格式");
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

  step("完成");
  return { added, unchanged: plan.unchanged, errors };
}

/** 沒有產生任何執行意圖時，印出前幾列待執行資料的關鍵欄位，協助判斷是哪個規則擋掉。 */
function logEmptyScheduleHints(texts: string[][], values: CellValue[][], schedule: SourceSchedule): void {
  const c = schedule.columns;
  const s = schedule.stats;
  log("[提示] 沒有產生任何待同步的執行意圖，可能原因：");
  if (s.dataRows === s.blankRows) log("  - 標題列以下全是空白。");
  if (s.completedRows > 0 && s.activeRows === 0) log(`  - 所有 ${s.completedRows} 列的「${SOURCE_HEADERS.result}」都有值，程式視為已結案而略過。`);
  if (s.missingChangeIdRows > 0) log(`  - ${s.missingChangeIdRows} 列「${SOURCE_HEADERS.changeId}」空白。`);
  if (s.invalidDateEntries > 0) {
    log(`  - ${s.invalidDateEntries} 筆「${SOURCE_HEADERS.changeDate}」無法辨識。支援格式：yyyy/m/d 或 yyyy-m-d（西元四位數年），或 Excel 日期值；民國年（如 115/11/20）不支援。`);
  }
  let shown = 0;
  for (let i = 1; i < texts.length && shown < LOG_SAMPLE_LIMIT; i++) {
    const row = texts[i];
    if (row.every((v) => !v.trim())) continue;
    shown++;
    log(`  第 ${i + 1} 列：${SOURCE_HEADERS.changeId}=「${row[c.changeId]}」 ${SOURCE_HEADERS.result}=「${row[c.result]}」 ` +
      `${SOURCE_HEADERS.changeDate}=「${row[c.changeDate]}」(值 ${JSON.stringify(values[i][c.changeDate])}) ` +
      `${SOURCE_HEADERS.schedule}=「${row[c.schedule].replace(/\r?\n/g, " ⏎ ")}」`);
  }
}
