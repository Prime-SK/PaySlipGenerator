import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from n_payslip import generate_payslips


class PaySlipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PaySlip Generator 2026")
        self.root.geometry("900x500")
        self.root.minsize(900, 500)
        self.root.resizable(False, False)
        self.root.configure(bg="#e8edf5")

        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Select an Excel file to begin")

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
        style.configure("Pill.TLabel", background="#1d4ed8", foreground="#ffffff", font=("Segoe UI Semibold", 9), padding=(10, 4))

        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 11),
            padding=(18, 10),
            foreground="#ffffff",
            background="#0d6efd",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#0b5ed7"), ("disabled", "#93c5fd")],
            foreground=[("disabled", "#f1f5f9")],
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(11, 7),
            foreground="#111827",
            background="#e7ebf2",
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#d9dee7")])

    def _build_layout(self):
        app = ttk.Frame(self.root, style="App.TFrame", padding=18)
        app.pack(fill="both", expand=True)

        hero = ttk.Frame(app, style="Hero.TFrame", padding=(22, 20))
        hero.pack(fill="x")

        hero_left = ttk.Frame(hero, style="Hero.TFrame")
        hero_left.pack(side="left", fill="x", expand=True)

        ttk.Label(hero_left, text="PaySlip Generator", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            hero_left,
            text="Generate polished employee payslips from Excel with a clean, guided workflow.",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(5, 0))

        ttk.Label(hero, text="Shakila Thathsara 2026", style="Pill.TLabel").pack(side="right", anchor="ne", padx=(16, 0), pady=(4, 0))

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

        footer = ttk.Frame(left_panel, style="Card.TFrame")
        footer.pack(fill="x", pady=(16, 0))
        ttk.Label(footer, text="Shakila Thathsara 2026", style="Brand.TLabel").pack(anchor="e")

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
        self.root.update_idletasks()

        try:
            count, date_range = generate_payslips(excel, output)
            self.status_text.set(f"Done: {count} payslips generated")
            messagebox.showinfo(
                "Success",
                f"Generated {count} payslips.\nPeriod: {date_range}\n\nSaved to:\n{output}",
            )
        except Exception as exc:
            self.status_text.set("Failed")
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
