import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from n_payslip import generate_payslips

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

def load_config():
    """Load registration number from config.json"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {'reg_no': ''}
    return {'reg_no': ''}

def save_config(config):
    """Save registration number to config.json"""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except IOError as e:
        print(f"Warning: Could not save config: {e}")


class PaySlipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ReliaPro Manpower - PaySlip Generator")
        self.root.geometry("980x620")
        self.root.minsize(980, 620)
        self.root.resizable(False, False)
        self.root.configure(bg="#e8edf5")

        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Select an Excel file to begin")
        self.output_text = tk.StringVar(value="Run output will appear here after generation.")
        
        self.config = load_config()

        self._configure_style()
        self._build_layout()

    def _configure_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("App.TFrame", background="#e8edf5")
        style.configure("Hero.TFrame", background="#0f172a")
        style.configure("Card.TFrame", background="#f9fbfd")
        style.configure("HeaderTitle.TLabel", background="#0f172a", foreground="#ffffff", font=("Segoe UI Semibold", 22))
        style.configure("HeaderSub.TLabel", background="#0f172a", foreground="#dbe4f0", font=("Segoe UI", 10))
        style.configure("FieldLabel.TLabel", background="#f9fbfd", foreground="#111827", font=("Segoe UI Semibold", 10))
        style.configure("Helper.TLabel", background="#f9fbfd", foreground="#5b6472", font=("Segoe UI", 9))
        style.configure("Path.TEntry", fieldbackground="#ffffff", bordercolor="#d1d5db", foreground="#111827", padding=(10, 8))
        style.configure("Status.TLabel", background="#f9fbfd", foreground="#0f172a", font=("Segoe UI Semibold", 10))
        style.configure("Brand.TLabel", background="#f9fbfd", foreground="#6b7280", font=("Segoe UI", 9, "italic"))
        style.configure("PanelTitle.TLabel", background="#f9fbfd", foreground="#111827", font=("Segoe UI Semibold", 10))
        style.configure("Pill.TLabel", background="#1d4ed8", foreground="#ffffff", font=("Segoe UI Semibold", 9), padding=(10, 4))

        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 11),
            padding=(18, 10),
            foreground="#ffffff",
            background="#0d6efd",
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#0b5ed7"), ("disabled", "#93c5fd"), ("!disabled", "#0d6efd")],
            foreground=[("disabled", "#f1f5f9"), ("!disabled", "#ffffff")],
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(11, 7),
            foreground="#111827",
            background="#e7ebf2",
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#d9dee7"), ("!disabled", "#e7ebf2")],
            foreground=[("!disabled", "#111827")],
        )
        
        style.configure(
            "Settings.TButton",
            font=("Segoe UI Semibold", 9),
            padding=(12, 6),
            foreground="#ffffff",
            background="#1f2937",
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "Settings.TButton",
            background=[("active", "#374151"), ("!disabled", "#1f2937")],
            foreground=[("!disabled", "#ffffff")],
        )

    def _build_layout(self):
        app = ttk.Frame(self.root, style="App.TFrame", padding=18)
        app.pack(fill="both", expand=True)

        hero = ttk.Frame(app, style="Hero.TFrame", padding=(22, 20))
        hero.pack(fill="x")

        hero_left = ttk.Frame(hero, style="Hero.TFrame")
        hero_left.pack(side="left", fill="x", expand=True)

        ttk.Label(hero_left, text="ReliaPro Manpower - PaySlip Generator", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            hero_left,
            text="Generate polished employee payslips from Excel with a clean, guided workflow.",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(5, 0))

        hero_right = ttk.Frame(hero, style="Hero.TFrame")
        hero_right.pack(side="right", anchor="ne", padx=(16, 0))
        
        self.settings_btn = ttk.Button(hero_right, text="Settings", style="Settings.TButton", command=self.show_settings)
        self.settings_btn.pack(side="left", padx=(0, 12))
        
        ttk.Label(hero_right, text="Shakila Thathsara", style="Pill.TLabel").pack(side="left")

        body = ttk.Frame(app, style="App.TFrame")
        body.pack(fill="both", expand=True, pady=(16, 0))

        left_panel = ttk.Frame(body, style="Card.TFrame", padding=20)
        left_panel.pack(fill="both", expand=True)

        top_accent = tk.Frame(left_panel, bg="#2563eb", height=4)
        top_accent.pack(fill="x", pady=(0, 16))

        ttk.Label(left_panel, text="Source Excel (.xlsx)", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(left_panel, text="Choose the workbook that contains the employee sheet.", style="Helper.TLabel").pack(anchor="w", pady=(2, 8))
        row1 = ttk.Frame(left_panel, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 14))
        self.excel_entry = ttk.Entry(row1, textvariable=self.excel_path, style="Path.TEntry")
        self.excel_entry.pack(side="left", fill="x", expand=True)
        self.pick_excel_btn = ttk.Button(row1, text="Browse", style="Secondary.TButton", command=self.browse_excel)
        self.pick_excel_btn.pack(side="left", padx=(10, 0))

        ttk.Label(left_panel, text="Destination PDF", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(left_panel, text="Pick where the finished payslips should be saved.", style="Helper.TLabel").pack(anchor="w", pady=(2, 8))
        row2 = ttk.Frame(left_panel, style="Card.TFrame")
        row2.pack(fill="x", pady=(0, 18))
        self.output_entry = ttk.Entry(row2, textvariable=self.output_path, style="Path.TEntry")
        self.output_entry.pack(side="left", fill="x", expand=True)
        self.pick_output_btn = ttk.Button(row2, text="Save As", style="Secondary.TButton", command=self.choose_output)
        self.pick_output_btn.pack(side="left", padx=(10, 0))

        action_bar = ttk.Frame(left_panel, style="Card.TFrame")
        action_bar.pack(fill="x", pady=(10, 0))

        status_block = ttk.Frame(action_bar, style="Card.TFrame")
        status_block.pack(side="left", fill="x", expand=True)
        ttk.Label(status_block, textvariable=self.status_text, style="Status.TLabel").pack(anchor="w")
        ttk.Label(
            status_block,
            text="The generate button unlocks after you select a workbook.",
            style="Helper.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        self.generate_btn = ttk.Button(action_bar, text="Generate Payslips", style="Primary.TButton", command=self.generate)
        self.generate_btn.pack(side="right")

        self._set_initial_state()

        output_panel = ttk.Frame(left_panel, style="Card.TFrame")
        output_panel.pack(fill="both", expand=True, pady=(16, 0))
        ttk.Label(output_panel, text="Output Summary", style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(
            output_panel,
            text="Rows converted, validation issues, and row-level details are shown below.",
            style="Helper.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        output_wrap = ttk.Frame(output_panel, style="Card.TFrame")
        output_wrap.pack(fill="both", expand=True)

        self.output_box = tk.Text(
            output_wrap,
            height=14,
            wrap="word",
            bg="#ffffff",
            fg="#0f172a",
            relief="solid",
            borderwidth=1,
            font=("Consolas", 10),
        )
        self.output_box.pack(side="left", fill="both", expand=True)

        output_scroll = ttk.Scrollbar(output_wrap, orient="vertical", command=self.output_box.yview)
        output_scroll.pack(side="right", fill="y")
        self.output_box.configure(yscrollcommand=output_scroll.set)
        self._set_output_panel(self.output_text.get())

        footer = ttk.Frame(left_panel, style="Card.TFrame")
        footer.pack(fill="x", pady=(16, 0))
        ttk.Label(footer, text="Shakila Thathsara", style="Brand.TLabel").pack(anchor="e")

    def _set_output_panel(self, text):
        self.output_box.config(state="normal")
        self.output_box.delete("1.0", tk.END)
        self.output_box.insert(tk.END, text)
        self.output_box.config(state="disabled")

    def _format_run_summary(self, summary):
        lines = [
            "Run complete",
            "",
            f"Source file: {summary.get('excel_file', '-')}",
            f"Output PDF: {summary.get('output_pdf', '-')}",
            f"Date range: {summary.get('date_range', '-')}",
            f"Rows scanned: {summary.get('rows_scanned', 0)}",
            f"Candidate rows: {summary.get('candidate_rows', 0)}",
            f"Payslips generated: {summary.get('employees_processed', 0)}",
            f"Validation issues: {len(summary.get('issues', []))}",
        ]

        report_path = summary.get('validation_report_path')
        if report_path:
            lines.append(f"Validation report: {report_path}")

        issues = summary.get('issues', [])
        if issues:
            lines.append("")
            lines.append("Issues")
            lines.append("------")
            for issue in issues:
                lines.append(
                    f"Row {issue['row']} | {issue['field']} | {issue['reason']} | raw={issue['raw_value']}"
                )

        return "\n".join(lines)

    def show_settings(self):
        """Open settings dialog for registration number"""
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
        ttk.Label(
            frame,
            text="Enter the company registration number to display on payslips.",
            style="Helper.TLabel"
        ).pack(anchor="w", pady=(2, 8))
        
        reg_entry = ttk.Entry(frame, style="Path.TEntry", width=30)
        reg_entry.pack(fill="x", pady=(0, 16))
        reg_entry.insert(0, current_reg)
        reg_entry.focus_set()
        
        button_frame = ttk.Frame(frame, style="Card.TFrame")
        button_frame.pack(fill="x", pady=(8, 0))
        
        def save_settings():
            new_reg = reg_entry.get().strip()
            if not new_reg:
                messagebox.showwarning("Invalid Input", "Registration number cannot be empty.")
                return
            self.config['reg_no'] = new_reg
            save_config(self.config)
            messagebox.showinfo("Success", "Registration number updated successfully.")
            dialog.destroy()
        
        ttk.Button(button_frame, text="Cancel", style="Secondary.TButton", command=dialog.destroy).pack(side="right")
        ttk.Button(button_frame, text="Save Settings", style="Primary.TButton", command=save_settings).pack(side="right", padx=(0, 8))
    
    def _set_initial_state(self):
        self.generate_btn.config(state="disabled")

    def _set_ready_state(self):
        self.status_text.set("Ready to generate")
        self.generate_btn.config(state="normal")

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select source Excel file",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            self.excel_path.set(path)
            if not self.output_path.get().strip():
                base_name = os.path.splitext(os.path.basename(path))[0]
                default_pdf = os.path.join(os.path.dirname(path), f"{base_name}_payslips.pdf")
                self.output_path.set(default_pdf)
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

    def generate(self):
        excel = self.excel_path.get().strip()
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

        self.pick_excel_btn.config(state="disabled")
        self.pick_output_btn.config(state="disabled")
        self.generate_btn.config(state="disabled")
        self.status_text.set("Generating payslips")
        self._set_output_panel("Running validation and PDF generation...")
        self.root.update_idletasks()

        try:
            reg_no = self.config.get('reg_no', 'NOT SET')
            summary = generate_payslips(excel, output, write_validation_report=False, reg_no=reg_no)
            count = summary.get('employees_processed', 0)
            issues_count = len(summary.get('issues', []))
            self.status_text.set(f"Done: {count} payslips generated")
            self._set_output_panel(self._format_run_summary(summary))

            messagebox.showinfo(
                "Success",
                f"Generated {count} payslips.\n"
                f"Validation issues: {issues_count}\n"
                f"Period: {summary.get('date_range', '-')}\n\n"
                f"Saved to:\n{output}",
            )
        except Exception as exc:
            self.status_text.set("Failed")
            self._set_output_panel(f"Run failed:\n{exc}")
            messagebox.showerror("Error", str(exc))
        finally:
            self.pick_excel_btn.config(state="normal")
            self.pick_output_btn.config(state="normal")
            if self.excel_path.get().strip():
                self._set_ready_state()


def main():
    root = tk.Tk()
    PaySlipApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
