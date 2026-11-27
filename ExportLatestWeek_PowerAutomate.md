# Export latest week workbook

## Purpose

`ExportLatestWeek.ts` is a standalone Office Script for the existing source workbook. Run it only on a newly created copy, never on the maintained source file. It keeps the first worksheet, its header row, and every data row whose column A value has the numerically largest valid `W<number>` week. Every other used row below the header is deleted in the copy.

The script uses the first worksheet by position, not by name. The source workbook must have the `數位平台部` sheet as its first worksheet and the used range must begin at A1. It does not create or reformat a worksheet; keeping the original workbook copy preserves workbook-level and sheet-level formatting, widths, tables, and other workbook contents.

Accepted week values are case-insensitive `W` followed by a positive integer, optionally surrounded by whitespace (for example `W01`, `w47`, `W 1` is not accepted). Comparison is numeric, so `W9` is less than `W10`. Output uses a minimum two-digit week label. Blank, malformed, and non-week values are deleted from the export and counted as `invalidWeekRowsDeleted`. If there is no valid week, the script throws before modifying the sheet.

## Power Automate flow

Create an instant cloud flow (manual trigger) for an initial controlled rollout. A manual trigger is preferable to a scheduled flow until the workbook location, run concurrency, and import process have been verified.

1. **Manually trigger a flow**: no inputs required.
2. **Get file content** (OneDrive for Business or SharePoint): select the maintained source `.xlsx` workbook.
3. **Create file** in a staging/export folder, using the source file content. Name it with a collision-resistant timestamp, for example `LatestWeek_yyyyMMdd_HHmmss.xlsx`. Do not overwrite the source or an existing export.
4. **Run script** (Excel Online (Business)): select the newly created copy as the workbook and `ExportLatestWeek` as the script. Store scripts in the OneDrive/SharePoint location supported by the chosen Run script action and tenant.
5. **Condition**: check the Run script result's `keptRows` is greater than zero. The script already fails when no valid week exists; this check protects against flow wiring mistakes.
6. **Move file** (or create the final file from the staging copy) into the import-ready folder only after the script succeeds and the condition passes. Include `latestWeek` in the final filename if your connector/action supports using the returned result dynamically; for example `W47_yyyyMMdd_HHmmss.xlsx`.
7. **Failure handling**: configure a failure scope (run-after: failed/timed out) to notify the operator and leave the copy in a clearly named staging/error folder. Never route a failed or zero-row result to the import-ready folder.

For the first test, use a non-production copy and verify the exported workbook in Excel plus the target system's import. Confirm that the first worksheet is `數位平台部`, the header/format are intact, only the largest numeric week remains, and other worksheets (if any) are acceptable to the import system. A copied workbook is a 1-to-1 workbook clone; this intentionally retains other sheets as well. If the receiving system interprets all sheets, its behavior must be checked before production use.

## Operational notes

- Keep the source workbook closed to conflicting edits while the flow copies it, or accept that the export represents the saved state at copy time.
- Set flow concurrency to 1 to avoid simultaneous runs and Excel file locks during the copy/script/move sequence.
- The script returns `worksheet`, `latestWeek`, `keptRows`, `deletedRows`, and `invalidWeekRowsDeleted` for flow logging.
- Excel Online and Power Automate connectors can hold locks briefly after script execution. If moving the workbook immediately fails, add a short retry/delay or use the connector's supported file operation pattern.
- If your tenant uses a SharePoint library for scripts, use the matching **Run script from SharePoint library** action; otherwise use **Run script** for OneDrive scripts.
