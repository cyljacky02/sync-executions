/**
 * Run against a copy of the source workbook. Keeps the latest W<number> rows
 * on the first worksheet and deletes every other used row below the header.
 */
type ExportResult = {
  worksheet: string;
  latestWeek: string;
  keptRows: number;
  deletedRows: number;
  invalidWeekRowsDeleted: number;
};

function main(workbook: ExcelScript.Workbook): ExportResult {
  const worksheets = workbook.getWorksheets();
  if (worksheets.length === 0) throw new Error("活頁簿沒有工作表。");

  const sheet = worksheets[0];
  const usedRange = sheet.getUsedRange(true);
  if (!usedRange) throw new Error(`工作表「${sheet.getName()}」沒有資料。`);
  if (usedRange.getRowIndex() !== 0 || usedRange.getColumnIndex() !== 0) {
    throw new Error(`工作表「${sheet.getName()}」的標題必須從 A1 開始，目前範圍為 ${usedRange.getAddress()}。`);
  }

  const rowCount = usedRange.getRowCount();
  if (rowCount < 2) throw new Error(`工作表「${sheet.getName()}」沒有標題以外的資料列。`);

  const weekValues = sheet.getRangeByIndexes(1, 0, rowCount - 1, 1).getTexts();
  const weeks: (number | null)[] = [];
  let latestWeekNumber = -1;

  weekValues.forEach((row) => {
    const text = row[0].trim();
    const match = text.match(/^W(\d+)$/i);
    const weekNumber = match ? Number(match[1]) : NaN;
    if (!Number.isSafeInteger(weekNumber) || weekNumber < 1) {
      weeks.push(null);
      return;
    }
    weeks.push(weekNumber);
    if (weekNumber > latestWeekNumber) latestWeekNumber = weekNumber;
  });

  if (latestWeekNumber < 1) {
    throw new Error("A 欄找不到有效週次（格式需為 W 加正整數，例如 W01、W47）。未修改工作表。");
  }

  const latestWeek = `W${String(latestWeekNumber).padStart(2, "0")}`;
  const deleteRuns: { startRow: number; count: number; invalidCount: number }[] = [];
  let keptRows = 0;
  let invalidWeekRowsDeleted = 0;

  // Build contiguous deletion runs and execute from bottom to top so row indexes stay valid.
  let runStart = -1;
  let runCount = 0;
  let runInvalidCount = 0;
  for (let i = 0; i < weeks.length; i++) {
    const shouldKeep = weeks[i] === latestWeekNumber;
    if (shouldKeep) {
      keptRows++;
      if (runCount > 0) {
        deleteRuns.push({ startRow: runStart, count: runCount, invalidCount: runInvalidCount });
        runStart = -1;
        runCount = 0;
        runInvalidCount = 0;
      }
    } else {
      if (runStart < 0) runStart = i + 2;
      runCount++;
      if (weeks[i] === null) runInvalidCount++;
    }
  }
  if (runCount > 0) {
    deleteRuns.push({ startRow: runStart, count: runCount, invalidCount: runInvalidCount });
  }

  let deletedRows = 0;
  for (let i = deleteRuns.length - 1; i >= 0; i--) {
    const run = deleteRuns[i];
    sheet.getRange(`A${run.startRow}:A${run.startRow + run.count - 1}`)
      .getEntireRow()
      .delete(ExcelScript.DeleteShiftDirection.up);
    deletedRows += run.count;
    invalidWeekRowsDeleted += run.invalidCount;
  }

  return {
    worksheet: sheet.getName(),
    latestWeek,
    keptRows,
    deletedRows,
    invalidWeekRowsDeleted
  };
}
