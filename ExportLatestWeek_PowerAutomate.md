# Export latest week workbook

## Purpose

`ExportLatestWeek.ts` is a standalone Office Script for the existing source workbook. Run it only on a newly created copy, never on the maintained source file. It keeps the first worksheet, its header row, and every data row whose column A value has the numerically largest valid `W<number>` week. Every other used row below the header is deleted in the copy.

The script uses the first worksheet by position, not by name. The source workbook must have the `數位平台部` sheet as its first worksheet and the used range must begin at A1. It does not create or reformat a worksheet; keeping the original workbook copy preserves workbook-level and sheet-level formatting, widths, tables, and other workbook contents.

Accepted week values are case-insensitive `W` followed by a positive integer, optionally surrounded by whitespace (for example `W01`, `w47`, `W 1` is not accepted). Comparison is numeric, so `W9` is less than `W10`. Output uses a minimum two-digit week label. Blank, malformed, and non-week values are deleted from the export and counted as `invalidWeekRowsDeleted`. If there is no valid week, the script throws before modifying the sheet.

## Power Automate flow

Create an instant cloud flow (manual trigger) for the initial rollout. After the flow is verified, the trigger can be replaced by a schedule or another approved trigger.

### Pipeline

```text
Manually trigger a flow
  -> OneDrive for Business: Copy file using path
  -> Excel Online (Business): Run script
  -> Condition: result.keptRows > 0
  -> OneDrive for Business: Move or rename a file
  -> import-ready folder
```

This pipeline copies the workbook first, then runs the script only on the copy. The original workbook is never edited. The copy action is preferable to `Get file content` plus `Create file` because it clones the `.xlsx` directly and avoids rebuilding the binary file through the flow.

### 1. Trigger

Use **Manually trigger a flow** with no inputs for the first version. Configure trigger concurrency to `1` after saving the flow so two exports cannot run against the same source at the same time.

### 2. Copy the source workbook

Add **OneDrive for Business -> Copy file using path**. If the action is shown as **Copy file**, use its `File` identifier input instead.

Set:

- **File Path**: the full path of the maintained source workbook, for example `/SyncExecution.xlsx`.
- **Destination File Path**: a temporary path containing a unique timestamp. Use the `fx` expression below and replace `/Export-Staging/` with the actual folder:

```text
concat('/Export-Staging/LatestWeek_',convertTimeZone(utcNow(),'UTC','Taipei Standard Time','yyyyMMdd_HHmmss'),'_',substring(guid(),0,8),'.xlsx')
```

- **Overwrite**: `No`.

The copy action output identifies the newly created file. Keep this output for the next steps. If `Run script` cannot accept that output directly, add **OneDrive for Business -> Get file metadata using path**, use the same destination path, and pass its `Id`/`Identifier` to `Run script`.

If the maintained workbook is in Teams or SharePoint rather than personal OneDrive, use the equivalent **SharePoint -> Copy file** action and the SharePoint version of file metadata. Do not mix OneDrive and SharePoint locations in the same copy action.

### 3. Run the Office Script on the copy

Add **Excel Online (Business) -> Run script**. Do not add `Get table`, `Create table`, `List rows present in a table`, `Get a row`, or `Delete a row` actions; this script works directly on the first worksheet and its used range.

Set:

- **Location**: `OneDrive for Business`.
- **Document Library**: `OneDrive`.
- **File**: the `Id`/`Identifier` of the copied workbook, not the source workbook.
- **Script**: `ExportLatestWeek`.

There are no table name, range, or script input parameters. The script returns `worksheet`, `latestWeek`, `keptRows`, `deletedRows`, and `invalidWeekRowsDeleted`.

### 4. Validate the result

Add a **Condition** immediately after `Run script`:

```text
@greater(int(outputs('Run_script')?['body']?['result']?['keptRows']), 0)
```

Replace `Run_script` with the internal name of the Run script action. In the Power Automate designer, selecting the returned `keptRows` dynamic content is safer than typing the path if the connector exposes it differently.

The script already fails when no valid week is found. This condition is an additional guard against an incorrectly configured flow. Only the `Yes` branch may move the file to the import-ready folder.

### 5. Rename and move the completed export

In the `Yes` branch, add **OneDrive for Business -> Move or rename a file**.

- **File**: the copied workbook's `Id`/`Identifier`.
- **Destination File Path**: use the returned `latestWeek` in the final filename. Example expression:

```text
concat('/Import-Ready/',outputs('Run_script')?['body']?['result']?['latestWeek'],'_LatestWeek_',convertTimeZone(utcNow(),'UTC','Taipei Standard Time','yyyyMMdd_HHmmss'),'.xlsx')
```

This produces a filename such as:

```text
/Import-Ready/W47_LatestWeek_20260320_173012.xlsx
```

Replace `Run_script` with the actual internal Action name. If dynamic content exposes `latestWeek`, insert that token in the expression rather than manually typing the result path. Set overwrite to `No` unless the import process explicitly requires replacement of an existing file.

### 6. Failure handling

Add a parallel failure scope or a final notification action configured with **run after** for `has failed`, `has timed out`, and `is skipped`. On failure, leave the copied workbook in `/Export-Staging/` or move it to an error folder. Do not move a failed or zero-row result to `/Import-Ready/`.

A short delay or retry may be needed before the move/rename because Excel Online can briefly hold a file lock after `Run script`.

For the first test, use a non-production copy and verify the exported workbook in Excel plus the target system's import. Confirm that the first worksheet is `數位平台部`, the header/format are intact, only the largest numeric week remains, and other worksheets (if any) are acceptable to the import system. A copied workbook intentionally retains other sheets as well; verify that the receiving system will not import unintended sheets.


## Operational notes

- Keep the source workbook closed to conflicting edits while the flow copies it, or accept that the export represents the saved state at copy time.
- Set flow concurrency to 1 to avoid simultaneous runs and Excel file locks during the copy/script/move sequence.
- The script returns `worksheet`, `latestWeek`, `keptRows`, `deletedRows`, and `invalidWeekRowsDeleted` for flow logging.
- Excel Online and Power Automate connectors can hold locks briefly after script execution. If moving the workbook immediately fails, add a short retry/delay or use the connector's supported file operation pattern.
- If your tenant uses a SharePoint library for scripts, use the matching **Run script from SharePoint library** action; otherwise use **Run script** for OneDrive scripts.
