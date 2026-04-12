# PaySlip Generator

A professional payslip generation tool for ReliaPro Manpower that converts Excel employee data into formatted PDF payslips.

## Features

- 📊 Excel to PDF conversion with professional formatting
- ✅ Automatic data validation with detailed error reporting
- 📝 Markdown validation reports alongside PDFs
- ⚙️ Configurable registration number
- 🎯 Real-time validation diagnostics in the UI
- 🔒 Safe numeric parsing and data integrity checks

## Quick Start

### Using the Executable

1. Extract `dist/PaySlipGen.exe` to your desired location
2. Double-click `PaySlipGen.exe` to launch the application
3. Click **⚙ Settings** to configure your company registration number (required before first use)
4. Select your Excel file (source payroll data)
5. Choose where to save the PDF output
6. Click **Generate Payslips**
7. Review the validation report for any data quality issues

### From Source (Development)

Open PowerShell in the project folder:

```powershell
cd D:\Projects\PaySlipGen

# Install dependencies
pip install openpyxl reportlab pyinstaller

# Run the UI
python payslip_ui.py
```

## Building the Executable

After modifying the UI (`payslip_ui.py`) or business logic (`n_payslip.py`), rebuild the executable:

```powershell
cd D:\Projects\PaySlipGen

# Clean rebuild with minimal file size
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PaySlipGen payslip_ui.py
```

The new executable will be created at: `dist\PaySlipGen.exe`

**Note:** Close `PaySlipGen.exe` if it's already running before rebuild.

### Build Command Breakdown

| Flag | Purpose |
|------|---------|
| `--noconfirm` | Skip confirmation prompts |
| `--clean` | Remove previous build artifacts |
| `--onefile` | Create a single executable (vs. folder) |
| `--windowed` | Hide console window (GUI-only) |
| `--name PaySlipGen` | Output executable name |
| `payslip_ui.py` | Entry point (UI module) |

## Configuration

The application stores the registration number in `config.json`:

```json
{
  "reg_no": "YOUR_COMPANY_REG_NUMBER"
}
```

You can also:
1. Click **⚙ Settings** button in the UI to update the registration number
2. Edit `config.json` directly with a text editor

The registration number will appear on all generated payslips.

## Command-Line Usage

The core generator can also be run directly from PowerShell:

```powershell
python n_payslip.py <input.xlsx> <output.pdf>
```

Example:

```powershell
python n_payslip.py ".\employee_data.xlsx" ".\payslips.pdf"
```

## Excel File Format

Your Excel file must contain:
- **Column headers** identifying employee data (Employee Name, Gross Pay, Net Salary, Account Number, etc.)
- **Date range** (format: `YYYY.MM.DD - YYYY.MM.DD` or `YYYY-MM-DD - YYYY-MM-DD`) somewhere in the top rows
- **Data rows** with employee information

See `111.xlsx` or `222.xlsx` for examples.

## Output Files

For each generation, the system creates:

1. **PDF file** - Formatted payslips (one page per employee)
2. **Markdown validation report** - Issues found during generation

Example: `payslips.pdf` creates `payslips.pdf.validation_report.md`

## Validation & Issue Reporting

The system validates:
- Employee numbers (must be integers)
- Numeric fields (salary, allowances, deductions)
- Required fields (name, gross pay, net salary)
- Data type consistency

Issues are logged with:
- Row number
- Field name
- Reason for the issue
- Raw value from Excel

## Project Structure

```
PaySlipGen/
├── payslip_ui.py           # Tkinter GUI application
├── n_payslip.py            # Core PDF generation & validation logic
├── config.json             # User settings (registration number)
├── dist/
│   └── PaySlipGen.exe     # Compiled executable
├── build/                  # Temporary build files (PyInstaller)
└── README.md              # This file
```

## Troubleshooting

### "Registration number cannot be empty"
- Click **⚙ Settings** and enter your company's registration number

### PDF generated but validation report shows issues
- Check the `.md` report file for details
- Verify column names match expected format
- Ensure date range is present in the Excel file

### Executable won't start
- Ensure Windows Defender or antivirus allows the `.exe`
- Verify `config.json` is in the same folder as `PaySlipGen.exe`
- Run from Command Prompt to see error messages

### "Could not find a valid header row"
- Verify the Excel file contains columns like:
  - Employee Name
  - Gross Pay
  - Net Salary
  - Basic Salary
  - Bank / Branch / Account Number

### PyInstaller build fails
- Close any running `PaySlipGen.exe` first
- Ensure all dependencies are installed: `pip install openpyxl reportlab pyinstaller`
- Try running: `pip install --upgrade pyinstaller`

## Requirements

- **Windows 10 or later** (Windows 11 recommended)
- **Python 3.13** (for development)
- Python packages:
  - `openpyxl` – Excel reading
  - `reportlab` – PDF generation
  - `pyinstaller` – Executable packaging

## Development Notes

### Key Functions

**payslip_ui.py**
- `PaySlipApp` – Main GUI class
- `load_config()` / `save_config()` – Configuration file management
- `show_settings()` – Registration number configuration dialog

**n_payslip.py**
- `generate_payslips()` – Main entry point for PDF generation
- `extract_employees()` – Parse Excel with validation
- `draw_payslip()` – ReportLab canvas drawing
- `parse_int_like()` / `parse_float_like()` – Safe numeric conversion

### Code Quality

All code maintains:
- Type safety for numeric conversions
- Row-level validation issue tracking
- Clean separation of UI and business logic
- Backward compatibility with existing Excel formats

## Branding

The application footer displays: **Shakila Thathsara**

Payslips include ReliaPro Manpower branding and the configured registration number.

## Version History

**Current (2026-04-08)**
- ✅ Settings button for registration number configuration
- ✅ Config file storage (`config.json`)
- ✅ Full validation diagnostics
- ✅ Output panel showing run summary
- ✅ Executable (21.5 MB)

## Support

For issues or questions, refer to the validation report generated alongside each PDF for detailed error information.
