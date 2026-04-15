# Installation Guide

## Prerequisites

### Required
- **Windows 10 or 11** (x64)
- **Python 3.11+** — [Download from python.org](https://www.python.org/downloads/)
- **Git** — for cloning the repository

### Optional
- **Microsoft Word** — required for PDF preview generation
- **Microsoft Excel** — for opening generated `.xlsx` files

---

## Step-by-Step Installation

### 1. Clone the Repository

```powershell
git clone https://github.com/<your-username>/marinadoc.git
cd marinadoc
```

### 2. Create Virtual Environment

```powershell
python -m venv .venv
```

### 3. Activate Environment

```powershell
# PowerShell
.\.venv\Scripts\Activate.ps1

# If you get an execution policy error:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\.venv\Scripts\Activate.ps1
```

### 4. Install Dependencies

```powershell
pip install -r requirements.txt
```

This installs:
- `PySide6` — Desktop UI framework
- `python-docx` — Word document processing
- `openpyxl` — Excel file processing
- `pydantic` — Data validation
- `pywin32` — Windows integration (Word automation)
- `pytest` — Testing framework

### 5. Run the Application

```powershell
python main.py
```

The MarinaDoc window should appear.

---

## Running Tests

```powershell
pytest -v
```

---

## Configuration

Application settings are stored in `config/app_config.json`:

```json
{
  "app_name": "MarinaDoc",
  "window_title": "MarinaDoc - Document Generator",
  "paths": {
    "templates_dir": "templates",
    "output_dir": "output",
    "preview_dir": "output/preview"
  },
  "templates": {
    "act_word_template_v2": "blank_templates/act_template.docx",
    "ks2_template": "blank_templates/ks2_template.xlsx",
    "ks3_template": "blank_templates/ks3_template.xlsx"
  },
  "logging_config_path": "config/logging.json"
}
```

---

## Adding Templates

1. Place your Word (`.docx`) and Excel (`.xlsx`) templates in `templates/blank_templates/`
2. Update template paths in `config/app_config.json`
3. Use `{{PLACEHOLDER}}` syntax for text replacement
4. Use `{{ROW_TEMPLATE}}` in a row to mark it for cloning

### Supported Placeholders

| Placeholder | Description |
|-------------|-------------|
| `{{contract_number}}` | Contract number |
| `{{contract_date}}` | Contract date |
| `{{customer_name_full}}` | Customer organization name |
| `{{executor_name_full}}` | Executor organization name |
| `{{total_with_vat}}` | Total amount with VAT |
| `{{ROW_TEMPLATE}}` | Row cloning marker |

---

## Troubleshooting

### "ModuleNotFoundError: No module named 'PySide6'"
Make sure the virtual environment is activated: `.\.venv\Scripts\Activate.ps1`

### Preview not working
Microsoft Word must be installed for PDF preview generation. The application will still work without it — documents are generated but won't show PDF preview.

### Permission denied on .venv
Run PowerShell as Administrator or adjust execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## System Requirements

| Requirement | Minimum | Recommended |
|------------|---------|-------------|
| OS | Windows 10 | Windows 11 |
| RAM | 4 GB | 8 GB |
| Disk | 500 MB | 1 GB |
| Python | 3.11 | 3.12+ |
| MS Office | Not required | Word + Excel (for preview) |
