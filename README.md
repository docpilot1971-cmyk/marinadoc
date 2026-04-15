# 📄 MarinaDoc — Document Automation Platform

> **DEMO Version** — Public showcase of the document automation engine.  
> Production version with advanced parsing heuristics and full template support is available on request.

---

## What It Does

MarinaDoc is a **Windows desktop application** that automatically processes construction contracts and generates the complete set of associated documents:

- **Contract Parsing** — extracts structured data from `.docx` contracts (parties, dates, object info, estimate tables, totals)
- **Act Generation** — produces formatted Acts (`.docx`) with all required fields and signatures
- **KS-2 Reports** — generates standardized KS-2 forms (`.xlsx`) with estimate line items
- **KS-3 Reports** — generates standardized KS-3 forms (`.xlsx`) with work summaries

All document generation is **template-driven** — swap templates to adapt to different document formats without changing code.

---

## Business Value

| Problem | Solution |
|---------|----------|
| Manual data entry from contracts → errors, lost time | Automated extraction and validation |
| Document generation takes hours | One-click generation of full document set |
| Template updates require retraining | Template-driven architecture — change templates, not code |
| Inconsistent document formatting | Standardized output from validated templates |

---

## Key Features

- 🔍 **Smart Parsing** — Multi-strategy text extraction from `.docx` contracts (paragraphs, tables, mixed layouts)
- 🏗️ **Contract Classification** — Auto-detects contract type (Organization / Individual Entrepreneur) and table structure
- 📋 **Estimate Table Extraction** — Handles sectional and flat estimate layouts, detects sections and line items
- ✅ **Validation Pipeline** — Checks completeness of extracted data before generation
- 📝 **Template-Driven Generation** — Word (`.docx`) and Excel (`.xlsx`) document generation from configurable templates
- 🎨 **Professional UI** — Full-featured desktop interface with document preview and editing
- 🔒 **Offline & Private** — No cloud, no API keys, no data leaves the machine

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        UI Layer (PySide6)                    │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │  Main Window │  │   Edit Form  │  │ Document Preview │  │
│  └──────┬───────┘  └──────┬───────┘  └────────┬─────────┘  │
└─────────┼─────────────────┼───────────────────┼────────────┘
          │                 │                   │
┌─────────▼─────────────────▼───────────────────▼────────────┐
│                   AppController                             │
│           Orchestrates the full pipeline                    │
└──────────────────────────┬──────────────────────────────────┘
                           │
          ┌────────────────┼────────────────┐
          ▼                ▼                ▼
    ┌──────────┐    ┌──────────┐    ┌──────────────┐
    │  Reader  │───▶│ Parsers  │───▶│  Validators  │
    │  (.docx) │    │ (header, │    │ (completeness│
    │          │    │ parties, │    │   & integrity)│
    │          │    │ tables,  │    │              │
    │          │    │ totals)  │    │              │
    └──────────┘    └────┬─────┘    └──────┬───────┘
                         │                 │
                         ▼                 ▼
                   ┌──────────────────────────┐
                   │     Generators           │
                   │  ┌────┐ ┌────┐ ┌──────┐ │
                   │  │Act │ │KS-2│ │ KS-3 │ │
                   │  │.doc│ │.xls│ │ .xls │ │
                   │  └────┘ └────┘ └──────┘ │
                   └──────────────────────────┘
```

### Module Structure

| Module | Responsibility |
|--------|---------------|
| `app/core/` | Configuration, logging, application controller |
| `app/models/` | Data models (Party, Estimate, Document, enums) |
| `app/services/` | Business logic: readers, parsers, generators, template processors |
| `app/services/interfaces/` | Abstract interfaces for all service contracts |
| `app/services/stubs/` | Demo implementations (full versions available on request) |
| `app/ui/` | PySide6 desktop interface (main window, preview widget) |
| `config/` | Application configuration files |
| `templates/` | Document templates (demo versions in this repo) |
| `tests/` | Test suite |

---

## Quick Start

### Prerequisites

- **Python 3.11+**
- **Windows 10/11** (uses `pywin32` for Word/Excel preview)
- **Microsoft Word** (optional, for PDF preview generation)

### Installation

```powershell
# 1. Clone the repository
git clone https://github.com/<your-username>/marinadoc.git
cd marinadoc

# 2. Create a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the application
python main.py
```

### Demo Workflow

1. Launch the app → `python main.py`
2. Open a `.docx` contract file
3. View the parsed data in the editing form
4. Click **Generate** to create the full document set (Act + KS-2 + KS-3)
5. Preview generated documents or open in Word/Excel

---

## This Is a DEMO Version

This repository contains a **simplified public version** of MarinaDoc. Key differences from production:

| Feature | This Demo | Production |
|---------|-----------|------------|
| Document parsing | Demo extraction (returns sample data) | Advanced multi-strategy regex pipeline |
| Contract classification | Static classification | Dynamic ORG/IP detection, table mode analysis |
| Template support | Demo templates | Production-ready templates for all contract types |
| Estimate extraction | Fixed demo rows | Full sectional/flat table parsing with section detection |
| Party extraction | Static demo parties | INN/OGRN/KPP extraction, bank details, addresses |

**The full production version is available on request.** Contact us to discuss licensing or collaboration.

---

## Screenshots

> *[Add screenshots here]*  
> - Main window with contract loaded
> - Editing form with extracted data
> - Generated document preview
> - Output folder with generated files

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| UI | PySide6 (Qt for Python) |
| Word Processing | python-docx |
| Excel Processing | openpyxl |
| Data Validation | pydantic |
| Word Automation (preview) | pywin32 |
| Testing | pytest |

---

## Project Structure

```
marinadoc/
├── main.py                      # Application entry point
├── requirements.txt             # Python dependencies
├── .gitignore                   # Git ignore rules
├── README.md                    # This file
│
├── app/                         # Application source code
│   ├── core/                    # Configuration, logging, controller
│   ├── models/                  # Data models and enums
│   ├── services/                # Business logic
│   │   ├── interfaces/          # Abstract service contracts
│   │   └── stubs/               # Demo implementations
│   └── ui/                      # PySide6 interface
│
├── config/                      # Configuration files
│   ├── app_config.json
│   └── logging.json
│
├── resources/                   # Application resources
│   ├── app_icon.ico
│   └── app_cover.png
│
├── templates/                   # Document templates (demo)
│   └── blank_templates/
│
└── tests/                       # Test suite
```

---

## License

This project is provided as a **demo showcase**. The source code in this repository is for evaluation purposes only.

For production use, licensing, or custom integrations, please contact the author.

---

## Contact

- **Author**: Marina  
- **Location**: Russia  
- **Email**: *[add your email]*  
- **Telegram**: *[add your Telegram]*  

> Interested in the full version or a similar solution for your business? Reach out — happy to discuss your use case.
