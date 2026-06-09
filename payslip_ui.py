import os
import json
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from n_payslip import generate_payslips, get_payroll_sheet_names

ALL_WEEKS = "All Weeks"


def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(get_app_dir(), 'config.json')


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {'reg_no': ''}
    return {'reg_no': ''}


def save_config(config):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except IOError as e:
        print(f"Warning: Could not save config: {e}")


def _safe_sheet_suffix(sheet_name):
    """Convert a sheet name like '1st Week' to a safe filename part like '1st_week'."""
    return sheet_name.strip().lower().replace(" ", "_").replace("/", "_")


def _build_output_paths(base_output, sheet_names):
    """
    Given the user's chosen output path and a list of sheet names,
    return a dict mapping each sheet name to its output PDF path.

    Example:
        base_output = "C:/out/payslips.pdf"
        sheet_names = ["1st Week", "2nd Week"]
        ->  {"1st Week": "C:/out/payslips_1st_week.pdf",
             "2nd Week": "C:/out/payslips_2nd_week.pdf"}
    """
    base = base_output[:-4] if base_output.lower().endswith('.pdf') else base_output
    return {s: f"{base}_{_safe_sheet_suffix(s)}.pdf" for s in sheet_names}


class PaySlipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ReliaPro Manpower - PaySlip Generator")
        self.root.geometry("980x740")
        self.root.minsize(980, 740)
        self.root.resizable(False, False)
        self.root.configure(bg="#e8edf5")

        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_var  = tk.StringVar(value=ALL_WEEKS)
        self.status_text = tk.StringVar(value="Select an Excel file to begin")
        self.output_text = tk.StringVar(value="Run output will appear here after generation.")

        self.config = load_config()
        self._available_sheets = []   # payroll sheet names detected from the workbook

        self._configure_style()
        self._build_layout()

    # ── Styles ────────────────────────────────────────────────────────────────

    def _configure_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("App.TFrame",    background="#e8edf5")
        style.configure("Hero.TFrame",   background="#0f172a")
        style.configure("Card.TFrame",   background="#f9fbfd", relief="flat", borderwidth=1)

        style.configure("HeaderTitle.TLabel", background="#0f172a", foreground="#ffffff",
                        font=("Segoe UI Semibold", 22))
        style.configure("HeaderSub.TLabel",   background="#0f172a", foreground="#dbe4f0",
                        font=("Segoe UI", 10))
        style.configure("FieldLabel.TLabel",  background="#f9fbfd", foreground="#111827",
                        font=("Segoe UI Semibold", 10))
        style.configure("Helper.TLabel",      background="#f9fbfd", foreground="#5b6472",
                        font=("Segoe UI", 9))
        style.configure("Path.TEntry",        fieldbackground="#ffffff", bordercolor="#cfd8e3",
                        foreground="#111827", padding=(12, 9), relief="flat")
        style.configure("Status.TLabel",      background="#f9fbfd", foreground="#0f172a",
                        font=("Segoe UI Semibold", 10))
        style.configure("Brand.TLabel",       background="#f9fbfd", foreground="#6b7280",
                        font=("Segoe UI", 9, "italic"))
        style.configure("PanelTitle.TLabel",  background="#f9fbfd", foreground="#111827",
                        font=("Segoe UI Semibold", 10))
        style.configure("Footer.TLabel",      background="#e8edf5", foreground="#4b5563",
                        font=("Segoe UI Semibold", 9))

        style.configure("Primary.TButton",
                        font=("Segoe UI Semibold", 10), padding=(14, 9),
                        foreground="#ffffff", background="#0d6efd",
                        borderwidth=0, relief="flat")
        style.map("Primary.TButton",
                  background=[("active", "#0b5ed7"), ("disabled", "#93c5fd"), ("!disabled", "#0d6efd")],
                  foreground=[("disabled", "#f1f5f9"), ("!disabled", "#ffffff")])

        style.configure("Secondary.TButton",
                        font=("Segoe UI Semibold", 10), padding=(14, 9),
                        foreground="#111827", background="#e7ebf2",
                        borderwidth=0, relief="flat")
        style.map("Secondary.TButton",
                  background=[("active", "#d9dee7"), ("!disabled", "#e7ebf2")],
                  foreground=[("!disabled", "#111827")])

        style.configure("Settings.TButton",
                        font=("Segoe UI Semibold", 10), padding=(14, 9),
                        foreground="#ffffff", background="#334155",
                        borderwidth=0, relief="flat")
        style.map("Settings.TButton",
                  background=[("active", "#475569"), ("!disabled", "#334155")],
                  foreground=[("!disabled", "#ffffff")])

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_layout(self):
        app = ttk.Frame(self.root, style="App.TFrame", padding=18)
        app.pack(fill="both", expand=True)

        # Hero bar
        hero = ttk.Frame(app, style="Hero.TFrame", padding=(22, 20))
        hero.pack(fill="x")

        hero_left = ttk.Frame(hero, style="Hero.TFrame")
        hero_left.pack(side="left", fill="x", expand=True)
        ttk.Label(hero_left, text="ReliaPro Manpower - PaySlip Generator",
                  style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(hero_left, text="Generate employee payslips from Excel.",
                  style="HeaderSub.TLabel").pack(anchor="w", pady=(5, 0))

        hero_right = ttk.Frame(hero, style="Hero.TFrame")
        hero_right.pack(side="right", anchor="ne", padx=(16, 0))
        self.settings_btn = ttk.Button(hero_right, text="⚙ Reg No", width=13,
                                       style="Settings.TButton", command=self.show_settings)
        self.settings_btn.pack(side="right")

        # Main card
        body = ttk.Frame(app, style="App.TFrame")
        body.pack(fill="both", expand=True, pady=(14, 0))

        card = ttk.Frame(body, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)

        top_accent = tk.Frame(card, bg="#2563eb", height=4)
        top_accent.pack(fill="x", pady=(0, 12))

        # ── Source Excel ──────────────────────────────────────────────────────
        ttk.Label(card, text="Source Excel (.xlsx)", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(card, text="Choose the workbook that contains the employee sheet(s).",
                  style="Helper.TLabel").pack(anchor="w", pady=(2, 6))
        row1 = ttk.Frame(card, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 12))
        self.excel_entry = ttk.Entry(row1, textvariable=self.excel_path, style="Path.TEntry")
        self.excel_entry.pack(side="left", fill="x", expand=True)
        self.pick_excel_btn = ttk.Button(row1, text="📂 Browse", width=13,
                                         style="Secondary.TButton", command=self.browse_excel)
        self.pick_excel_btn.pack(side="left", padx=(10, 0))

        # ── Sheet selector ────────────────────────────────────────────────────
        ttk.Label(card, text="Week / Sheet", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(card,
                  text="Select a specific week or 'All Weeks' to generate separate PDFs for each.",
                  style="Helper.TLabel").pack(anchor="w", pady=(2, 6))
        sheet_row = ttk.Frame(card, style="Card.TFrame")
        sheet_row.pack(fill="x", pady=(0, 12))
        self.sheet_dropdown = ttk.Combobox(sheet_row, textvariable=self.sheet_var,
                                            state="disabled", font=("Segoe UI", 10), width=28)
        self.sheet_dropdown['values'] = [ALL_WEEKS]
        self.sheet_dropdown.pack(side="left")
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        # ── Destination PDF ───────────────────────────────────────────────────
        ttk.Label(card, text="Destination PDF", style="FieldLabel.TLabel").pack(anchor="w")
        self.pdf_helper_label = ttk.Label(
            card,
            text="Where to save the payslips. With 'All Weeks', this is used as a base name.",
            style="Helper.TLabel",
        )
        self.pdf_helper_label.pack(anchor="w", pady=(2, 6))
        row2 = ttk.Frame(card, style="Card.TFrame")
        row2.pack(fill="x", pady=(0, 12))
        self.output_entry = ttk.Entry(row2, textvariable=self.output_path, style="Path.TEntry")
        self.output_entry.pack(side="left", fill="x", expand=True)
        self.pick_output_btn = ttk.Button(row2, text="💾 Save As", width=13,
                                          style="Secondary.TButton", command=self.choose_output)
        self.pick_output_btn.pack(side="left", padx=(10, 0))

        # ── Action bar ────────────────────────────────────────────────────────
        action_bar = ttk.Frame(card, style="Card.TFrame")
        action_bar.pack(fill="x", pady=(6, 0))

        status_block = ttk.Frame(action_bar, style="Card.TFrame")
        status_block.pack(side="left", fill="x", expand=True)
        ttk.Label(status_block, textvariable=self.status_text, style="Status.TLabel").pack(anchor="w")
        ttk.Label(status_block,
                  text="The generate button unlocks after you select a workbook.",
                  style="Helper.TLabel").pack(anchor="w", pady=(4, 0))

        self.generate_btn = ttk.Button(action_bar, text="▶ Generate Payslips", width=20,
                                       style="Primary.TButton", command=self.generate)
        self.generate_btn.pack(side="right")
        self._set_initial_state()

        # ── Output log ────────────────────────────────────────────────────────
        output_panel = ttk.Frame(card, style="Card.TFrame")
        output_panel.pack(fill="both", expand=True, pady=(12, 0))
        ttk.Label(output_panel, text="Output Summary", style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(output_panel,
                  text="Rows converted, validation issues, and row-level details are shown below.",
                  style="Helper.TLabel").pack(anchor="w", pady=(2, 8))

        output_wrap = ttk.Frame(output_panel, style="Card.TFrame")
        output_wrap.pack(fill="both", expand=True)
        self.output_box = tk.Text(output_wrap, height=8, wrap="word",
                                  bg="#ffffff", fg="#0f172a", relief="flat",
                                  borderwidth=1, font=("Consolas", 10))
        self.output_box.pack(side="left", fill="both", expand=True)
        output_scroll = ttk.Scrollbar(output_wrap, orient="vertical",
                                      command=self.output_box.yview)
        output_scroll.pack(side="right", fill="y")
        self.output_box.configure(yscrollcommand=output_scroll.set)
        self._set_output_panel(self.output_text.get())

        # Footer
        footer = ttk.Frame(app, style="App.TFrame")
        footer.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Label(footer, text="Developed by Shakila Thathsara",
                  style="Footer.TLabel").pack(anchor="e")

    # ── Output panel helpers ──────────────────────────────────────────────────

    def _set_output_panel(self, text):
        self.output_box.config(state="normal")
        self.output_box.delete("1.0", tk.END)
        self.output_box.insert(tk.END, text)
        self.output_box.config(state="disabled")

    def _format_single_summary(self, summary):
        lines = [
            "Run complete",
            "",
            f"Sheet          : {summary.get('sheet_name', '-')}",
            f"Date range     : {summary.get('date_range', '-')}",
            f"Rows scanned   : {summary.get('rows_scanned', 0)}",
            f"Candidate rows : {summary.get('candidate_rows', 0)}",
            f"Generated      : {summary.get('employees_processed', 0)} payslips",
            f"Issues         : {len(summary.get('issues', []))}",
            f"Output PDF     : {summary.get('output_pdf', '-')}",
        ]
        issues = summary.get('issues', [])
        if issues:
            lines += ["", "Issues", "------"]
            for issue in issues:
                lines.append(
                    f"Row {issue['row']} | {issue['field']} | "
                    f"{issue['reason']} | raw={issue['raw_value']}"
                )
        return "\n".join(lines)

    def _format_multi_summary(self, results):
        """results: list of (sheet_name, summary) tuples."""
        total_slips  = sum(s.get('employees_processed', 0) for _, s in results)
        total_issues = sum(len(s.get('issues', []))         for _, s in results)
        lines = [
            f"All Weeks complete — {len(results)} sheet(s) processed",
            f"Total payslips : {total_slips}",
            f"Total issues   : {total_issues}",
            "",
        ]
        for sheet, summary in results:
            lines += [
                f"── {sheet} ──",
                f"  Date range : {summary.get('date_range', '-')}",
                f"  Generated  : {summary.get('employees_processed', 0)} payslips",
                f"  Issues     : {len(summary.get('issues', []))}",
                f"  Output     : {summary.get('output_pdf', '-')}",
                "",
            ]
            for issue in summary.get('issues', []):
                lines.append(
                    f"  Row {issue['row']} | {issue['field']} | "
                    f"{issue['reason']} | raw={issue['raw_value']}"
                )
        return "\n".join(lines)

    # ── State helpers ─────────────────────────────────────────────────────────

    def _set_initial_state(self):
        self.generate_btn.config(state="disabled")

    def _set_ready_state(self):
        self.status_text.set("Ready to generate")
        self.generate_btn.config(state="normal")

    def _lock_ui(self):
        self.pick_excel_btn.config(state="disabled")
        self.pick_output_btn.config(state="disabled")
        self.generate_btn.config(state="disabled")
        self.sheet_dropdown.config(state="disabled")

    def _unlock_ui(self):
        self.pick_excel_btn.config(state="normal")
        self.pick_output_btn.config(state="normal")
        self.sheet_dropdown.config(state="readonly")
        if self.excel_path.get().strip():
            self._set_ready_state()

    # ── Sheet dropdown ────────────────────────────────────────────────────────

    def _populate_sheet_dropdown(self, excel_path):
        """Detect payroll sheets and fill the dropdown."""
        try:
            sheets = get_payroll_sheet_names(excel_path)
        except Exception:
            sheets = []

        self._available_sheets = sheets
        choices = [ALL_WEEKS] + sheets

        self.sheet_dropdown['values'] = choices
        self.sheet_var.set(ALL_WEEKS)
        self.sheet_dropdown.config(state="readonly")

    def _on_sheet_changed(self, _event=None):
        """Update output path suggestion when the sheet selection changes."""
        excel = self.excel_path.get().strip()
        if not excel:
            return
        base_name = os.path.splitext(os.path.basename(excel))[0]
        folder    = os.path.dirname(excel)
        selected  = self.sheet_var.get()

        if selected == ALL_WEEKS:
            # Show generic base name — actual files will be e.g. base_1st_week.pdf
            self.output_path.set(os.path.join(folder, f"{base_name}_payslips.pdf"))
        else:
            safe = _safe_sheet_suffix(selected)
            self.output_path.set(os.path.join(folder, f"{base_name}_{safe}.pdf"))

    # ── File dialogs ──────────────────────────────────────────────────────────

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select source Excel file",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return

        self.excel_path.set(path)
        self._populate_sheet_dropdown(path)

        # Auto-fill output path
        base_name = os.path.splitext(os.path.basename(path))[0]
        default_pdf = os.path.join(os.path.dirname(path), f"{base_name}_payslips.pdf")
        self.output_path.set(default_pdf)

        sheet_count = len(self._available_sheets)
        label = f"{sheet_count} payroll sheet(s) found" if sheet_count else "No payroll sheets found"
        self.status_text.set(label)
        self._set_ready_state()

    def choose_output(self):
        initial_name = "payslips.pdf"
        if self.excel_path.get().strip():
            base_name = os.path.splitext(os.path.basename(self.excel_path.get().strip()))[0]
            initial_name = f"{base_name}_payslips.pdf"

        path = filedialog.asksaveasfilename(
            title="Select destination PDF",
            defaultextension=".pdf",
            initialfile=initial_name,
            filetypes=[("PDF file", "*.pdf")],
        )
        if path:
            self.output_path.set(path)
            if self.excel_path.get().strip():
                self._set_ready_state()

    # ── Generate ──────────────────────────────────────────────────────────────

    def generate(self):
        excel  = self.excel_path.get().strip()
        output = self.output_path.get().strip()

        if not excel:
            messagebox.showerror("Missing input", "Please select an Excel file.")
            return
        if not output:
            messagebox.showerror("Missing output", "Please select a destination PDF file.")
            return
        if not os.path.isfile(excel):
            messagebox.showerror("File not found", "Selected Excel file does not exist.")
            return
        if not excel.lower().endswith(".xlsx"):
            messagebox.showerror("Invalid file", "Please select a .xlsx file.")
            return

        selected = self.sheet_var.get()
        reg_no   = self.config.get('reg_no', 'NOT SET')

        self._lock_ui()
        self.status_text.set("Generating payslips…")
        self._set_output_panel("Running…")
        self.root.update_idletasks()

        try:
            if selected == ALL_WEEKS:
                self._generate_all_weeks(excel, output, reg_no)
            else:
                self._generate_single(excel, output, reg_no, selected)

        except Exception as exc:
            self.status_text.set("Failed")
            self._set_output_panel(f"Run failed:\n{exc}")
            messagebox.showerror("Error", str(exc))
        finally:
            self._unlock_ui()

    def _generate_single(self, excel, output, reg_no, sheet_name):
        summary = generate_payslips(
            excel, output,
            write_validation_report=False,
            reg_no=reg_no,
            sheet_name=sheet_name,
        )
        count = summary.get('employees_processed', 0)
        self.status_text.set(f"Done: {count} payslips generated")
        self._set_output_panel(self._format_single_summary(summary))
        self._show_single_success_dialog(summary, output)

    def _generate_all_weeks(self, excel, output, reg_no):
        sheets = self._available_sheets
        if not sheets:
            messagebox.showerror("No sheets", "No payroll sheets were detected in this workbook.")
            return

        output_map = _build_output_paths(output, sheets)
        results = []

        for i, sheet in enumerate(sheets, 1):
            self.status_text.set(f"Processing sheet {i}/{len(sheets)}: {sheet}…")
            self.root.update_idletasks()

            out = output_map[sheet]
            summary = generate_payslips(
                excel, out,
                write_validation_report=False,
                reg_no=reg_no,
                sheet_name=sheet,
            )
            results.append((sheet, summary))

        total = sum(s.get('employees_processed', 0) for _, s in results)
        self.status_text.set(f"Done: {total} payslips across {len(results)} sheet(s)")
        self._set_output_panel(self._format_multi_summary(results))
        self._show_multi_success_dialog(results, output)

    # ── Success dialogs ───────────────────────────────────────────────────────

    def _open_output_folder(self, output_pdf):
        abs_pdf = os.path.abspath(output_pdf)
        folder  = os.path.dirname(abs_pdf)
        if not os.path.isdir(folder):
            messagebox.showerror("Folder not found", f"Could not find folder:\n{folder}")
            return
        try:
            subprocess.Popen(f'explorer /select,"{abs_pdf}"')
        except Exception:
            try:
                os.startfile(folder)
            except Exception as exc:
                messagebox.showerror("Open Folder Failed", str(exc))

    def _show_single_success_dialog(self, summary, output_pdf):
        dialog = tk.Toplevel(self.root)
        dialog.title("Generation Complete")
        dialog.geometry("520x260")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg="#e8edf5")

        frame = ttk.Frame(dialog, style="Card.TFrame", padding=20)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ttk.Label(frame, text="Payslips Generated Successfully",
                  style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(frame,
                  text=(
                      f"Sheet          : {summary.get('sheet_name', '-')}\n"
                      f"Generated      : {summary.get('employees_processed', 0)} payslips\n"
                      f"Validation issues: {len(summary.get('issues', []))}\n"
                      f"Period         : {summary.get('date_range', '-')}\n\n"
                      f"Saved to:\n{output_pdf}"
                  ),
                  style="Helper.TLabel", justify="left").pack(anchor="w", pady=(8, 16))

        bf = ttk.Frame(frame, style="Card.TFrame")
        bf.pack(fill="x")
        ttk.Button(bf, text="Close", width=13, style="Secondary.TButton",
                   command=dialog.destroy).pack(side="right")
        ttk.Button(bf, text="📂 Open Folder", width=13, style="Primary.TButton",
                   command=lambda: self._open_output_folder(output_pdf)).pack(side="right", padx=(0, 8))

    def _show_multi_success_dialog(self, results, base_output):
        total  = sum(s.get('employees_processed', 0) for _, s in results)
        issues = sum(len(s.get('issues', []))         for _, s in results)

        dialog = tk.Toplevel(self.root)
        dialog.title("All Weeks Complete")
        dialog.geometry("560x340")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg="#e8edf5")

        frame = ttk.Frame(dialog, style="Card.TFrame", padding=20)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ttk.Label(frame, text="All Weeks Generated Successfully",
                  style="PanelTitle.TLabel").pack(anchor="w")

        # Per-sheet detail
        detail_lines = [
            f"Total payslips     : {total}",
            f"Total issues       : {issues}",
            f"Sheets processed   : {len(results)}",
            "",
        ]
        for sheet, summary in results:
            detail_lines.append(
                f"  {sheet:<14}  {summary.get('employees_processed', 0):>3} payslips"
                f"  |  {summary.get('date_range', 'N/A')}"
            )
        detail_lines += ["", "Files saved in:", os.path.dirname(os.path.abspath(base_output))]

        ttk.Label(frame, text="\n".join(detail_lines),
                  style="Helper.TLabel", justify="left").pack(anchor="w", pady=(8, 16))

        bf = ttk.Frame(frame, style="Card.TFrame")
        bf.pack(fill="x")
        ttk.Button(bf, text="Close", width=13, style="Secondary.TButton",
                   command=dialog.destroy).pack(side="right")
        ttk.Button(bf, text="📂 Open Folder", width=13, style="Primary.TButton",
                   command=lambda: self._open_output_folder(base_output)).pack(side="right", padx=(0, 8))

    # ── Settings ──────────────────────────────────────────────────────────────

    def show_settings(self):
        current_reg = self.config.get('reg_no', '')

        dialog = tk.Toplevel(self.root)
        dialog.title("Settings")
        dialog.geometry("430x230")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg="#e8edf5")

        frame = ttk.Frame(dialog, style="Card.TFrame", padding=20)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ttk.Label(frame, text="Registration Number", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(frame,
                  text="Enter the company registration number to display on payslips.",
                  style="Helper.TLabel").pack(anchor="w", pady=(2, 8))

        reg_entry = ttk.Entry(frame, style="Path.TEntry", width=30)
        reg_entry.pack(fill="x", pady=(0, 16))
        reg_entry.insert(0, current_reg)
        reg_entry.focus_set()

        bf = ttk.Frame(frame, style="Card.TFrame")
        bf.pack(fill="x", pady=(8, 0))

        def save_settings():
            new_reg = reg_entry.get().strip()
            if not new_reg:
                messagebox.showwarning("Invalid Input", "Registration number cannot be empty.")
                return
            self.config['reg_no'] = new_reg
            save_config(self.config)
            messagebox.showinfo("Success", "Registration number updated successfully.")
            dialog.destroy()

        ttk.Button(bf, text="✖ Cancel", width=13, style="Secondary.TButton",
                   command=dialog.destroy).pack(side="right")
        ttk.Button(bf, text="✔ Save", width=13, style="Primary.TButton",
                   command=save_settings).pack(side="right", padx=(0, 8))


def main():
    root = tk.Tk()
    PaySlipApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()