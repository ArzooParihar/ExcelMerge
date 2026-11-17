import os
import re
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_TITLE = "Generic Excel Merger & Validator"
SUPPORTED = {".xlsx", ".xlsm"}

def clean_header(value):
    """Keep the visible header text, removing only outer whitespace."""
    if value is None:
        return ""
    return str(value).strip()

def header_key(value):
    """Normalize case and ignore every whitespace/newline difference."""
    text = clean_header(value).casefold()
    return re.sub(r"\s+", "", text)

def find_header_row(ws, base_headers=None, search_rows=5):
    """Find the most likely header row among the first five rows.

    For input sheets, prefer the row matching the largest number of Base headers;
    ties are resolved by the number of populated cells, then earliest row.
    For the Base sheet, use the row with the most populated cells.
    """
    base_keys = {header_key(h) for h in (base_headers or []) if header_key(h)}
    best = None
    max_row = min(search_rows, ws.max_row or search_rows)
    for row_num in range(1, max_row + 1):
        values = [clean_header(c.value) for c in ws[row_num]]
        while values and values[-1] == "":
            values.pop()
        nonempty = sum(bool(v) for v in values)
        if not nonempty:
            continue
        overlap = sum(1 for v in values if header_key(v) and header_key(v) in base_keys)
        score = (overlap if base_keys else 0, nonempty, -row_num)
        if best is None or score > best[0]:
            best = (score, row_num, values)
    if best is None:
        return 1, []
    return best[1], best[2]

def get_headers(ws, base_headers=None):
    """Return (header_row_number, headers) after scanning the first five rows."""
    return find_header_row(ws, base_headers=base_headers)

def read_sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names

def choose_sheet(parent, title, sheet_names):
    """Skip the dialog when the workbook contains only one worksheet."""
    if len(sheet_names) == 1:
        return sheet_names[0]
    result = {"sheet": None}
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.grab_set()
    win.resizable(False, False)
    frame = ttk.Frame(win, padding=14)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="Select the worksheet to use:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))
    selected = tk.StringVar(value=sheet_names[0] if sheet_names else "")
    combo = ttk.Combobox(frame, textvariable=selected, values=sheet_names, state="readonly", width=38)
    combo.pack(fill="x", pady=(0, 12))
    buttons = ttk.Frame(frame)
    buttons.pack(fill="x")
    def ok():
        result["sheet"] = selected.get()
        win.destroy()
    def cancel():
        win.destroy()
    ttk.Button(buttons, text="Select", command=ok).pack(side="right")
    ttk.Button(buttons, text="Cancel", command=cancel).pack(side="right", padx=(0, 8))
    win.protocol("WM_DELETE_WINDOW", cancel)
    parent.wait_window(win)
    return result["sheet"]

def is_empty_row(row):
    return not any(value is not None and not (isinstance(value, str) and value.strip() == "") for value in row)

def scan_data_rows(ws, header_row=1):
    """
    Count non-empty data rows, but only accept the consecutive block after the header row.
    The first completely empty row ends the data block. Any later non-empty rows are skipped.
    Returns (usable_rows, skipped_rows, total_nonempty_rows).
    """
    usable = skipped = total = 0
    gap_found = False
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if is_empty_row(row):
            gap_found = True
            continue
        total += 1
        if gap_found:
            skipped += 1
        else:
            usable += 1
    return usable, skipped, total

def header_problems(base_headers, candidate_headers):
    problems = []
    if not base_headers:
        return ["Base worksheet has no headers in the first five rows."]

    if any(h == "" for h in base_headers):
        blanks = [i + 1 for i, h in enumerate(base_headers) if h == ""]
        problems.append("Base Excel has blank header cell(s) inside the header range, column(s): " + ", ".join(map(str, blanks)))

    base_keys = [header_key(h) for h in base_headers]
    candidate_keys = [header_key(h) for h in candidate_headers]

    base_counts = Counter(k for k in base_keys if k)
    base_dupes = list(dict.fromkeys(base_headers[i] for i, k in enumerate(base_keys) if k and base_counts[k] > 1))
    if base_dupes:
        problems.append("Base Excel has duplicate header(s), after header normalization: " + ", ".join(repr(x) for x in base_dupes))

    if any(h == "" for h in candidate_headers):
        blanks = [i + 1 for i, h in enumerate(candidate_headers) if h == ""]
        problems.append("Blank header cell(s) inside the header range, column(s): " + ", ".join(map(str, blanks)))

    candidate_counts = Counter(k for k in candidate_keys if k)
    dupes = list(dict.fromkeys(candidate_headers[i] for i, k in enumerate(candidate_keys) if k and candidate_counts[k] > 1))
    if dupes:
        problems.append("Duplicate header(s), after header normalization: " + ", ".join(repr(x) for x in dupes))

    optional_base_keys = {header_key(h) for h in base_headers if header_key(h) == "remark"}
    base_required_keys = set(base_keys) - optional_base_keys
    candidate_set = set(candidate_keys)
    missing = [h for h in base_headers if header_key(h) in base_required_keys and header_key(h) not in candidate_set]
    additional = [h for h in candidate_headers if h and header_key(h) not in set(base_keys)]
    if missing:
        problems.append("Missing column(s): " + ", ".join(repr(x) for x in missing))
    if additional:
        problems.append("Additional column(s): " + ", ".join(repr(x) for x in additional))
    return problems

def build_column_validation(base_headers, candidate_headers):
    """Return per-Base-column status plus workbook-level header issues."""
    base_keys = [header_key(h) for h in base_headers]
    candidate_keys = [header_key(h) for h in candidate_headers]
    counts = Counter(k for k in candidate_keys if k)
    statuses = {}
    other = []
    for header, key in zip(base_headers, base_keys):
        matches = [i for i, candidate_key in enumerate(candidate_keys) if candidate_key == key and key]
        if len(matches) > 1:
            statuses[key] = "ERROR: Duplicate header in input"
        elif len(matches) == 0:
            if key == "remark":
                statuses[key] = "✓ (optional; blank)"
            else:
                statuses[key] = "ERROR: Missing column"
        else:
            statuses[key] = "✓"
    blank_positions = [str(i + 1) for i, h in enumerate(candidate_headers) if not h]
    if blank_positions:
        other.append("Blank header at input column(s): " + ", ".join(blank_positions))
    extras = [h for h in candidate_headers if h and header_key(h) not in set(base_keys)]
    if extras:
        other.append("Additional column(s): " + ", ".join(extras))
    # Duplicate headers that do not map to a Base column are still reported.
    for key, count in counts.items():
        if count > 1 and key not in set(base_keys):
            names = [candidate_headers[i] for i, k in enumerate(candidate_keys) if k == key]
            other.append("Duplicate extra header(s): " + ", ".join(names))
    return statuses, other

class MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("850x610")
        self.root.minsize(760, 540)
        self.base_path = None
        self.base_sheet = None
        self.files = []  # dicts: path, sheet
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents"))
        self.status = tk.StringVar(value="First select a Base Excel file.")
        self.build_ui()

    def build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=APP_TITLE, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(main, text="Use any workbook layout as the template. Column order is matched by header name, not position.",
                  wraplength=800).pack(anchor="w", pady=(3, 14))

        base_box = ttk.LabelFrame(main, text="1. Base Excel (defines the required columns)", padding=10)
        base_box.pack(fill="x", pady=(0, 10))
        row = ttk.Frame(base_box); row.pack(fill="x")
        self.base_label = ttk.Label(row, text="No base file selected", width=75)
        self.base_label.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse Base Excel", command=self.select_base).pack(side="right")
        sheetrow = ttk.Frame(base_box); sheetrow.pack(fill="x", pady=(8, 0))
        ttk.Label(sheetrow, text="Base worksheet:").pack(side="left")
        self.base_sheet_var = tk.StringVar()
        self.base_combo = ttk.Combobox(sheetrow, textvariable=self.base_sheet_var, state="readonly", width=35)
        self.base_combo.pack(side="left", padx=8)
        self.base_combo.bind("<<ComboboxSelected>>", self.on_base_sheet_change)
        ttk.Label(sheetrow, text="Header is auto-detected within the first 5 rows.", foreground="#555").pack(side="left", padx=8)
        self.base_info = ttk.Label(base_box, text="Base columns: not loaded yet", wraplength=790, justify="left")
        self.base_info.pack(anchor="w", pady=(8, 0))

        input_box = ttk.LabelFrame(main, text="2. Excel files to merge", padding=10)
        input_box.pack(fill="both", expand=True, pady=(0, 10))
        toolbar = ttk.Frame(input_box); toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="Add Excel Files", command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Clear List", command=self.clear_files).pack(side="left")
        columns = ("file", "sheet", "rows")
        self.tree = ttk.Treeview(input_box, columns=columns, show="headings", height=9, selectmode="extended")
        self.tree.heading("file", text="Excel File")
        self.tree.heading("sheet", text="Selected Worksheet")
        self.tree.heading("rows", text="Data Rows")
        self.tree.column("file", width=430)
        self.tree.column("sheet", width=180)
        self.tree.column("rows", width=90, anchor="center")
        self.tree.pack(fill="both", expand=True)
        ttk.Label(input_box, text="When you add a workbook, the software asks which worksheet to validate and merge.",
                  foreground="#555").pack(anchor="w", pady=(7, 0))

        out_box = ttk.LabelFrame(main, text="3. Output location", padding=10)
        out_box.pack(fill="x", pady=(0, 10))
        outrow = ttk.Frame(out_box); outrow.pack(fill="x")
        ttk.Entry(outrow, textvariable=self.output_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(outrow, text="Browse", command=self.select_output).pack(side="right", padx=(8, 0))

        bottom = ttk.Frame(main); bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status, wraplength=580).pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="Validate & Merge", command=self.merge, style="Accent.TButton").pack(side="right")

        note = ("Rules: header matching is case-insensitive; outer spaces are ignored; column order may differ. "
                "Remark is optional if present in Base Excel; trailing blank header cells are ignored. "
                "Internal blank/duplicate/additional headers reject a file. Data after the first fully empty row is skipped and reported.")
        ttk.Label(main, text=note, wraplength=810, foreground="#555").pack(anchor="w", pady=(10, 0))

    def select_base(self):
        path = filedialog.askopenfilename(title="Select Base Excel", filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            names = read_sheet_names(path)
        except Exception as e:
            messagebox.showerror("Cannot open Base Excel", str(e)); return
        sheet = choose_sheet(self.root, "Choose Base Worksheet", names)
        if not sheet:
            return
        try:
            wb = load_workbook(path, read_only=True, data_only=False)
            ws = wb[sheet]
            header_row, headers = get_headers(ws)
            usable_rows, skipped_rows, total_rows = scan_data_rows(ws, header_row)
            wb.close()
        except Exception as e:
            messagebox.showerror("Cannot read Base Excel", str(e)); return
        self.base_path, self.base_sheet = path, sheet
        self.base_header_row = header_row
        self.base_label.config(text=path)
        self.base_sheet_var.set(sheet)
        self.base_combo["values"] = names
        header_text = ", ".join(f"{i+1}. {h}" for i, h in enumerate(headers))
        self.base_info.config(text=f"Base columns: {len(headers)} | Rows to merge: {usable_rows} | Rows after first blank row (skipped): {skipped_rows}\\nColumn names: {header_text}")
        self.status.set("Base Excel selected. Its data rows will also be included in the final Master Excel.")

    def on_base_sheet_change(self, _event=None):
        self.base_sheet = self.base_sheet_var.get()

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Select Excel files", filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        for path in paths:
            normalized = os.path.normcase(os.path.realpath(path))
            if self.base_path and normalized == os.path.normcase(os.path.realpath(self.base_path)):
                messagebox.showerror("Duplicate/Base file", f"This is already selected as the Base Excel and cannot be added again:\n{path}")
                continue
            duplicate = next((x for x in self.files if os.path.normcase(os.path.realpath(x["path"])) == normalized), None)
            if duplicate:
                messagebox.showerror("Duplicate file", f"This Excel file has already been added:\n{path}")
                continue
            try:
                names = read_sheet_names(path)
                sheet = choose_sheet(self.root, f"Choose worksheet: {Path(path).name}", names)
                if sheet:
                    wb = load_workbook(path, read_only=True, data_only=False)
                    ws = wb[sheet]
                    base_headers_for_scan = []
                    if self.base_path and self.base_sheet:
                        base_wb = load_workbook(self.base_path, read_only=True, data_only=False)
                        base_ws = base_wb[self.base_sheet]
                        _, base_headers_for_scan = get_headers(base_ws)
                        base_wb.close()
                    header_row, candidate_headers_for_scan = get_headers(ws, base_headers_for_scan)
                    usable_rows, skipped_rows, total_rows = scan_data_rows(ws, header_row)
                    wb.close()
                    item = {"path": path, "sheet": sheet, "header_row": header_row, "rows": usable_rows,
                            "skipped_rows": skipped_rows, "total_rows": total_rows}
                    self.files.append(item)
                    row_display = f"{usable_rows} / {total_rows}" + (f" (skip {skipped_rows})" if skipped_rows else "")
                    self.tree.insert("", "end", iid=str(len(self.files)-1),
                                     values=(Path(path).name, sheet, row_display))
            except Exception as e:
                messagebox.showerror("Cannot open workbook", f"{path}\n\n{e}")
        self.status.set(f"{len(self.files)} input file(s) selected.")

    def remove_selected(self):
        selected = self.tree.selection()
        if not selected: return
        remove_indices = sorted((int(i) for i in selected), reverse=True)
        for i in remove_indices:
            self.files.pop(i)
        self.refresh_tree()

    def clear_files(self):
        self.files.clear()
        self.refresh_tree()
        self.status.set("Input file list cleared.")

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, item in enumerate(self.files):
            row_display = f"{item.get('rows', 0)} / {item.get('total_rows', item.get('rows', 0))}"
            if item.get("skipped_rows", 0):
                row_display += f" (skip {item['skipped_rows']})"
            self.tree.insert("", "end", iid=str(i), values=(Path(item["path"]).name, item["sheet"], row_display))

    def select_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path: self.output_dir.set(path)

    def merge(self):
        if not self.base_path or not self.base_sheet:
            messagebox.showwarning("Base Excel required", "Please select the Base Excel and its worksheet first."); return
        if not self.files:
            messagebox.showwarning("Input files required", "Please add at least one Excel file."); return
        out_dir = self.output_dir.get().strip()
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning("Output folder required", "Please select an existing output folder."); return

        try:
            base_wb = load_workbook(self.base_path, read_only=True, data_only=False)
            base_ws = base_wb[self.base_sheet]
            base_header_row, base_headers = get_headers(base_ws)
            base_wb.close()
        except Exception as e:
            messagebox.showerror("Base Excel error", str(e)); return

        base_issues = header_problems(base_headers, base_headers)
        # The self-comparison above checks base blank/duplicate headers; remove comparison-derived differences (none).
        if base_issues:
            messagebox.showerror("Invalid Base Excel", "Please fix the Base Excel first:\n\n" + "\n".join(base_issues))
            return

        output_wb = Workbook(write_only=True)
        output_ws = output_wb.create_sheet("Merged Data")
        output_ws.append(base_headers)
        report = []
        rejected_files = []
        accepted = 0
        total_rows = 0

        # Include Base Excel data first, stopping at the first completely empty row.
        try:
            base_data_wb = load_workbook(self.base_path, read_only=True, data_only=False)
            base_data_ws = base_data_wb[self.base_sheet]
            base_index = {header_key(h): i for i, h in enumerate(base_headers)}
            base_usable, base_skipped, base_total = scan_data_rows(base_data_ws, base_header_row)
            base_gap_seen = False
            base_merged = 0
            for row in base_data_ws.iter_rows(min_row=base_header_row + 1, values_only=True):
                if is_empty_row(row):
                    base_gap_seen = True
                    continue
                if base_gap_seen:
                    continue
                output_ws.append([row[base_index[header_key(h)]] if base_index[header_key(h)] < len(row) else None for h in base_headers])
                total_rows += 1
                base_merged += 1
            base_data_wb.close()
            accepted += 1
            detail = f"Base worksheet: {self.base_sheet}; header at Row {base_header_row}; {base_merged} row(s) included"
            if base_skipped:
                detail += f"; WARNING: {base_total} non-empty data row(s) found, {base_skipped} row(s) after first completely empty row were not merged"
            report.append({"file": Path(self.base_path).name, "sheet": self.base_sheet, "status": "ACCEPTED (BASE)", "detail": detail, "columns": {header_key(h): "✓" for h in base_headers}, "other": ""})
        except Exception as e:
            messagebox.showerror("Base Excel read error", str(e))
            return

        for item in self.files:
            path, sheet = item["path"], item["sheet"]
            try:
                wb = load_workbook(path, read_only=True, data_only=False)
                if sheet not in wb.sheetnames:
                    report.append({"file": Path(path).name, "sheet": sheet, "status": "REJECTED", "detail": f"Worksheet not found: {sheet}", "columns": {header_key(h): "ERROR: Worksheet not found" for h in base_headers}, "other": "Worksheet not found"})
                    rejected_files.append(path)
                    wb.close()
                    continue
                ws = wb[sheet]
                header_row, candidate_headers = get_headers(ws, base_headers)
                issues = header_problems(base_headers, candidate_headers)
                if issues:
                    col_status, other_issues = build_column_validation(base_headers, candidate_headers)
                    report.append({"file": Path(path).name, "sheet": sheet, "status": "REJECTED", "detail": " | ".join(issues), "columns": col_status, "other": " | ".join(other_issues)})
                    rejected_files.append(path)
                    wb.close()
                    continue

                index_by_header = {header_key(h): i for i, h in enumerate(candidate_headers)}
                optional_remark = next((h for h in base_headers if header_key(h) == "remark"), None)
                usable, skipped, total_found = scan_data_rows(ws, header_row)
                gap_seen = False
                merged_from_file = 0
                for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                    if is_empty_row(row):
                        gap_seen = True
                        continue
                    if gap_seen:
                        continue
                    mapped = []
                    for header in base_headers:
                        key = header_key(header)
                        if key in index_by_header:
                            idx = index_by_header[key]
                            mapped.append(row[idx] if idx < len(row) else None)
                        elif optional_remark and header == optional_remark:
                            mapped.append(None)
                        else:
                            mapped.append(None)
                    output_ws.append(mapped)
                    total_rows += 1
                    merged_from_file += 1

                wb.close()
                accepted += 1
                detail = f"Headers matched; worksheet: {sheet}; header at Row {header_row}; {merged_from_file} row(s) merged"
                if optional_remark and header_key(optional_remark) not in index_by_header:
                    detail += f"; optional '{optional_remark}' column absent, filled with blank values"
                if skipped:
                    detail += f"; WARNING: {total_found} non-empty data row(s) found, {skipped} row(s) after first completely empty row were not merged"
                col_status, other_issues = build_column_validation(base_headers, candidate_headers)
                report.append({"file": Path(path).name, "sheet": sheet, "status": "ACCEPTED", "detail": detail, "columns": col_status, "other": " | ".join(other_issues)})
            except Exception as e:
                report.append({"file": Path(path).name, "sheet": sheet, "status": "REJECTED", "detail": f"Could not process workbook: {e}", "columns": {header_key(h): "ERROR: Processing failed" for h in base_headers}, "other": str(e)})
                rejected_files.append(path)

        stamp = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
        master_path = os.path.join(out_dir, f"Master_Merged_{stamp}.xlsx")
        report_path = os.path.join(out_dir, f"Validation_Report_{stamp}.xlsx")
        rejected_dir = os.path.join(out_dir, f"Rejected_Files_{stamp}")

        # Create a separate rejected-files folder for every merge run.
        os.makedirs(rejected_dir, exist_ok=True)
        copied_rejected = 0
        copy_errors = []
        for rejected_path in rejected_files:
            try:
                source_name = Path(rejected_path).name
                target = os.path.join(rejected_dir, source_name)
                stem, ext = os.path.splitext(source_name)
                suffix = 1
                while os.path.exists(target):
                    target = os.path.join(rejected_dir, f"{stem}_{suffix}{ext}")
                    suffix += 1
                shutil.copy2(rejected_path, target)
                copied_rejected += 1
            except Exception as copy_error:
                copy_errors.append(f"{Path(rejected_path).name}: {copy_error}")

        if accepted:
            try:
                output_wb.save(master_path)
            except Exception as e:
                messagebox.showerror("Save failed", str(e)); return
        else:
            messagebox.showwarning("No valid files", "No input file passed header validation. No Master Excel was created.")

        try:
            rw = Workbook()
            rws = rw.active
            rws.title = "Validation Report"
            report_headers = ["File Name", "Worksheet", "Overall Status"] + list(base_headers) + ["Other Issues", "Details"]
            rws.append(report_headers)
            for cell in rws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            green_fill = PatternFill("solid", fgColor="C6EFCE")
            red_fill = PatternFill("solid", fgColor="FFC7CE")
            for entry in report:
                values = [entry["file"], entry.get("sheet", ""), entry["status"]]
                for header in base_headers:
                    values.append(entry.get("columns", {}).get(header_key(header), "✓"))
                values.extend([entry.get("other", ""), entry.get("detail", "")])
                rws.append(values)
                current_row = rws.max_row
                for col_idx in range(4, 4 + len(base_headers)):
                    cell = rws.cell(current_row, col_idx)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    if str(cell.value).startswith("ERROR"):
                        cell.font = Font(color="9C0006", bold=True)
                        cell.fill = red_fill
                    else:
                        cell.font = Font(color="006100", bold=True)
                        cell.fill = green_fill
            rws.freeze_panes = "D2"
            rws.auto_filter.ref = rws.dimensions
            rws.column_dimensions["A"].width = 32
            rws.column_dimensions["B"].width = 20
            rws.column_dimensions["C"].width = 18
            for col_idx in range(4, 4 + len(base_headers)):
                rws.column_dimensions[get_column_letter(col_idx)].width = 24
            other_col = 4 + len(base_headers)
            detail_col = other_col + 1
            rws.column_dimensions[get_column_letter(other_col)].width = 42
            rws.column_dimensions[get_column_letter(detail_col)].width = 65
            for row in rws.iter_rows(min_row=2):
                row[2].font = Font(color="C00000" if row[2].value == "REJECTED" else "008000", bold=True)
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            if copy_errors:
                warning_row = ["Rejected-file copy warning", "", "WARNING"] + [""] * len(base_headers) + [" | ".join(copy_errors), ""]
                rws.append(warning_row)
            rw.save(report_path)
        except Exception as e:
            messagebox.showerror("Report save failed", str(e)); return

        added_accepted = accepted - 1  # exclude Base Excel from input-file count
        rejected_count = len(rejected_files)
        msg = (f"Excel merge completed successfully.\n\n"
               f"Accepted input files: {added_accepted}\n"
               f"Rejected input files: {rejected_count}\n"
               f"Rejected files copied: {copied_rejected}\n"
               f"Total data rows in final Master (Base included): {total_rows}\n\n"
               f"Validation report:\n{report_path}\n\n"
               f"Rejected files folder:\n{rejected_dir}")
        if accepted:
            msg += f"\n\nMaster Excel:\n{master_path}"
        if copy_errors:
            msg += "\n\nSome rejected files could not be copied. See Validation Report for details."
        self.status.set(f"Merge complete: {added_accepted} accepted, {rejected_count} rejected; {total_rows} final data rows. Rejected files: {rejected_dir}")
        self.root.bell()
        messagebox.showinfo("Merge Complete - Success", msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = MergerApp(root)
    root.mainloop()
