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
    """Normalize only surrounding whitespace. Matching remains case-sensitive."""
    if value is None:
        return ""
    return str(value).strip()

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

def get_headers(ws):
    """Read Row 1 and discard only blank header cells at the far right."""
    headers = [clean_header(c.value) for c in ws[1]]
    while headers and headers[-1] == "":
        headers.pop()
    return headers

def is_empty_row(row):
    return not any(value is not None and not (isinstance(value, str) and value.strip() == "") for value in row)

def scan_data_rows(ws):
    """
    Count non-empty data rows, but only accept the consecutive block starting at Row 2.
    The first completely empty row ends the data block. Any later non-empty rows are skipped.
    Returns (usable_rows, skipped_rows, total_nonempty_rows).
    """
    usable = skipped = total = 0
    gap_found = False
    for row in ws.iter_rows(min_row=2, values_only=True):
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
        return ["Base worksheet has no headers in Row 1."]
    if any(h == "" for h in base_headers):
        blanks = [i + 1 for i, h in enumerate(base_headers) if h == ""]
        problems.append("Base Excel has blank header cell(s) inside the header range, column(s): " + ", ".join(map(str, blanks)))
    base_dupes = [h for h, count in Counter(base_headers).items() if h and count > 1]
    if base_dupes:
        problems.append("Base Excel has duplicate header(s): " + ", ".join(repr(x) for x in base_dupes))
    if any(h == "" for h in candidate_headers):
        blanks = [i + 1 for i, h in enumerate(candidate_headers) if h == ""]
        problems.append("Blank header cell(s) inside the header range, column(s): " + ", ".join(map(str, blanks)))
    dupes = [h for h, count in Counter(candidate_headers).items() if h and count > 1]
    if dupes:
        problems.append("Duplicate header(s): " + ", ".join(repr(x) for x in dupes))

    # Remark is optional when it is a column in the Base Excel.
    optional_base_headers = {h for h in base_headers if h.lower() == "remark"}
    base_required = set(base_headers) - optional_base_headers
    candidate_set = set(candidate_headers)
    missing = [h for h in base_headers if h in base_required and h not in candidate_set]
    additional = [h for h in candidate_headers if h and h not in set(base_headers)]
    if missing:
        problems.append("Missing column(s): " + ", ".join(repr(x) for x in missing))
    if additional:
        problems.append("Additional column(s): " + ", ".join(repr(x) for x in additional))
    return problems

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
        ttk.Label(sheetrow, text="Headers are always read from Row 1.", foreground="#555").pack(side="left", padx=8)
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

        note = ("Rules: exact header names (case-sensitive; outer spaces are ignored); column order may differ. "
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
            headers = get_headers(ws)
            usable_rows, skipped_rows, total_rows = scan_data_rows(ws)
            wb.close()
        except Exception as e:
            messagebox.showerror("Cannot read Base Excel", str(e)); return
        self.base_path, self.base_sheet = path, sheet
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
                    usable_rows, skipped_rows, total_rows = scan_data_rows(ws)
                    wb.close()
                    item = {"path": path, "sheet": sheet, "rows": usable_rows,
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
            base_headers = get_headers(base_ws)
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
            base_index = {h: i for i, h in enumerate(base_headers)}
            base_usable, base_skipped, base_total = scan_data_rows(base_data_ws)
            base_gap_seen = False
            base_merged = 0
            for row in base_data_ws.iter_rows(min_row=2, values_only=True):
                if is_empty_row(row):
                    base_gap_seen = True
                    continue
                if base_gap_seen:
                    continue
                output_ws.append([row[base_index[h]] if base_index[h] < len(row) else None for h in base_headers])
                total_rows += 1
                base_merged += 1
            base_data_wb.close()
            accepted += 1
            detail = f"Base worksheet: {self.base_sheet}; {base_merged} row(s) included"
            if base_skipped:
                detail += f"; WARNING: {base_total} non-empty data row(s) found, {base_skipped} row(s) after first completely empty row were not merged"
            report.append((Path(self.base_path).name, "ACCEPTED (BASE)", detail))
        except Exception as e:
            messagebox.showerror("Base Excel read error", str(e))
            return

        for item in self.files:
            path, sheet = item["path"], item["sheet"]
            try:
                wb = load_workbook(path, read_only=True, data_only=False)
                if sheet not in wb.sheetnames:
                    report.append((Path(path).name, "REJECTED", f"Worksheet not found: {sheet}"))
                    rejected_files.append(path)
                    wb.close()
                    continue
                ws = wb[sheet]
                candidate_headers = get_headers(ws)
                issues = header_problems(base_headers, candidate_headers)
                if issues:
                    report.append((Path(path).name, "REJECTED", " | ".join(issues)))
                    rejected_files.append(path)
                    wb.close()
                    continue

                index_by_header = {h: i for i, h in enumerate(candidate_headers)}
                optional_remark = next((h for h in base_headers if h.lower() == "remark"), None)
                usable, skipped, total_found = scan_data_rows(ws)
                gap_seen = False
                merged_from_file = 0
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if is_empty_row(row):
                        gap_seen = True
                        continue
                    if gap_seen:
                        continue
                    mapped = []
                    for header in base_headers:
                        if header in index_by_header:
                            idx = index_by_header[header]
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
                detail = f"Headers matched; worksheet: {sheet}; {merged_from_file} row(s) merged"
                if optional_remark and optional_remark not in index_by_header:
                    detail += f"; optional '{optional_remark}' column absent, filled with blank values"
                if skipped:
                    detail += f"; WARNING: {total_found} non-empty data row(s) found, {skipped} row(s) after first completely empty row were not merged"
                report.append((Path(path).name, "ACCEPTED", detail))
            except Exception as e:
                report.append((Path(path).name, "REJECTED", f"Could not process workbook: {e}"))
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
            rws.append(["File Name", "Status", "Details"])
            for cell in rws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for entry in report:
                rws.append(list(entry))
            rws.freeze_panes = "A2"
            rws.auto_filter.ref = rws.dimensions
            for col, width in {"A": 34, "B": 14, "C": 110}.items():
                rws.column_dimensions[col].width = width
            for row in rws.iter_rows(min_row=2):
                if row[1].value == "REJECTED":
                    row[1].font = Font(color="C00000", bold=True)
                else:
                    row[1].font = Font(color="008000", bold=True)
                row[2].alignment = Alignment(wrap_text=True, vertical="top")
            if copy_errors:
                rws.append(["Rejected-file copy warning", "WARNING", " | ".join(copy_errors)])
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
