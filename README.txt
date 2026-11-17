# Generic Excel Merger & Validator

## Features
- Select any `.xlsx` or `.xlsm` workbook as the Base Excel.
- Select the Base worksheet; headers are auto-detected within the first five rows; rows above the detected header are ignored.
- Add multiple `.xlsx` / `.xlsm` input workbooks and choose a worksheet from each; the GUI displays each worksheet's data-row count.
- Match columns by header name, not by column position; the GUI displays the Base Excel's column count and header names.
- Input column order may differ; output follows the Base Excel column order.
- Reject a file if it has missing, additional, blank, or duplicate headers.
- Empty data cells are allowed and are not validated.
- Includes the Base Excel's data rows plus all accepted input files; displays the final merged data-row total in the GUI and completion message. Creates a Master Excel (when at least one added input file is valid) and a separate Validation Report Excel.
- Original files are not modified.

## Requirements
- Windows 10/11
- Python 3.10 or newer
- openpyxl

## Run on Windows
1. Install Python from https://www.python.org/downloads/ (enable "Add Python to PATH").
2. Open Command Prompt in this folder.
3. Run:
   `py -m pip install -r requirements.txt`
4. Run:
   `py generic_excel_merger.py`

## Header matching rule
Header matching ignores letter case and all whitespace differences, including extra/missing spaces, tabs, and line breaks anywhere in the header. Spaces around parentheses are also ignored. For example, `Employee Name`, `EMPLOYEE  NAME`, and `Employee\nName` match; `Date (Joining)` and `Date(Joining)` also match. Header row is auto-detected within the first five rows.

## Important
- The Base Excel defines the required columns AND its own data rows are included first in the final Master Excel.
- Only the selected worksheet from each input workbook is processed.
- Supported input formats in this first version: `.xlsx` and `.xlsm`.
- Password-protected workbooks and legacy `.xls` files are not supported in this version.


## V3 behavior updates
- If Base Excel has a `Remark` column, input files may omit it; missing Remark values are filled as blank in the Master.
- If a workbook has only one worksheet, it is selected automatically without showing a sheet-selection dialog.
- Trailing blank header cells to the right of the last named header are ignored. Blank header cells inside the header range are still rejected.
- Adding the same file more than once is blocked immediately with an error message. The Base Excel cannot also be added as an input.
- Data rows are read consecutively from the row immediately after the detected header until the first completely empty row. All rows after that first blank row are skipped and reported. The GUI displays usable/total non-empty rows and skipped count.

## V4 updates
- After processing, a clear Tkinter success popup appears with accepted/rejected file counts, final merged row count, and output paths. The app status line is also updated and the window gives an audible bell notification.
- Every merge run creates a uniquely named `Rejected_Files_YYYYMMDD_HHMMSS` subfolder inside the selected output folder.
- All input workbooks rejected during validation/processing are copied into that run's rejected-files folder. Original files are not moved or modified.
- If two rejected source files have the same filename, the copied filename gets a numeric suffix to prevent overwriting.
- The V4 completion popup also shows how many rejected files were successfully copied.


## V5 behavior update
- Header matching is case-insensitive. For example, `Name`, `NAME`, and `name` are treated as the same header.
- Leading/trailing spaces are ignored.
- Output column names and order continue to follow the Base Excel exactly.
- Headers differing only by letter case are duplicates and cause rejection.


## V6 updates
- Header matching ignores capitalization and all whitespace differences, including extra/missing spaces, tabs, and line breaks; spaces around parentheses are ignored.
- The app searches the first five rows of each selected worksheet for the header row. Rows above the detected header are ignored; data starts immediately below the header.
- Validation Report now has one separate validation column for every Base Excel header. A green check indicates a valid/matched column; a red error explains a missing or duplicate column. Additional/blank header issues are shown in a separate Other Issues column.
