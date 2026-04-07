# PaySlipGen

PaySlipGen is a Windows payslip generator that reads employee data from an Excel workbook and exports one PDF payslip per employee.

The project has two entry points:

- [n_payslip.py](n_payslip.py) for the core generation logic and command-line usage
- [payslip_ui.py](payslip_ui.py) for the desktop GUI

## Features

- Reads an `.xlsx` workbook and automatically detects the header row
- Extracts the payslip date range from the top rows of the sheet
- Generates a formatted PDF payslip for each employee
- Provides a simple Windows UI to choose the Excel file and destination PDF
- Can be packaged as a standalone `.exe` with PyInstaller

## Repository Layout

- [n_payslip.py](n_payslip.py) - core Excel-to-PDF generation script
- [payslip_ui.py](payslip_ui.py) - Tkinter desktop UI
- [PaySlipGen.spec](PaySlipGen.spec) - PyInstaller build file generated during packaging
- [dist/PaySlipGen.exe](dist/PaySlipGen.exe) - built executable output

## Requirements

- Windows 10 or Windows 11
- Python 3.13 or newer
- `pip`
- An Excel workbook in `.xlsx` format

Python packages used by the project:

- `openpyxl`
- `reportlab`
- `pyinstaller` for building the executable

## Quick Start

If you already have Python installed, you can install the dependencies and run the GUI directly.

### 1. Open PowerShell in the project folder

```powershell
cd D:\Projects\PaySlipGen
```

### 2. Create a virtual environment

```powershell
python -m venv .venv
```

### 3. Activate the virtual environment

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks script activation, run this once in the same window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate again.

### 4. Install dependencies

```powershell
pip install openpyxl reportlab pyinstaller
```

## How to Use the GUI

Run the desktop app with:

```powershell
python .\payslip_ui.py
```

In the window:

1. Click `Browse` and select an existing `.xlsx` file.
2. Click `Save As` and choose where to save the output PDF.
3. Click `Generate Payslips`.

The app will generate a PDF with one payslip per employee.

Branding shown in the UI footer:

- `Shakila Thathsara 2026`

## How to Run the Script from the Command Line

The core generator can also be run without the GUI.

### Basic usage

```powershell
python .\n_payslip.py "D:\path\to\source.xlsx" "D:\path\to\output.pdf"
```

### Example

```powershell
python .\n_payslip.py ".\Paid Labour salary 2026 . 02 . 28- 2026 .03.06.xlsx" new_payslip.pdf
```

If you omit the arguments, the script falls back to its default paths inside `n_payslip.py`.

## How the Date Range Is Extracted

The script scans the top rows of the worksheet and looks for date ranges in text such as:

```text
ATTENDANCE SHEET - Date:      2026.02.28 -2026.03.06
```

It supports common range styles like:

- `2026.02.28 -2026.03.06`
- `2026-02-28 - 2026-03-06`
- `2026/02/28 - 2026/03/06`

## How to Build the EXE

The project uses PyInstaller to package the GUI into a single Windows executable.

### 1. Install PyInstaller

```powershell
pip install pyinstaller
```

### 2. Build the executable

Run this from the project folder:

```powershell
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PaySlipGen .\payslip_ui.py
```

### 3. Find the output

When the build completes, the executable is created at:

```text
dist\PaySlipGen.exe
```

## Rebuilding After Changes

If you change the UI or the generator code, rebuild the executable with the same PyInstaller command.

If `dist\PaySlipGen.exe` is already open, close it first. Windows will lock the file and PyInstaller may fail to overwrite it.

## Troubleshooting

### The app says `Could not find a valid header row`

Make sure the workbook contains the employee table and the header row includes columns similar to:

- Employee Name
- Gross Pay
- Net Salary
- Basic Salary

### The date range shows `N/A`

Check that the workbook contains a date range in the top rows above the table headers.

### The executable build fails with a file lock error

Close any running copy of `PaySlipGen.exe` and rebuild.

### PowerShell blocks activation of the virtual environment

Run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Notes

- The project is optimized for Windows desktop use.
- The UI is intentionally simple and focused on file selection and PDF generation.
- If you want a custom app icon, a versioned installer, or an MSI package, those can be added later.

## Build Summary

The standard build flow is:

1. Install Python dependencies
2. Run the GUI with `python .\payslip_ui.py`
3. Validate the output PDF
4. Package the app with PyInstaller
5. Use `dist\PaySlipGen.exe`
