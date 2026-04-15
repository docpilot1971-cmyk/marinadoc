# Architecture Overview

## System Design

MarinaDoc follows a **layered architecture** with clear separation of concerns:

```
┌─────────────────────────────────────────────────────┐
│                  Presentation Layer                  │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────┐ │
│  │  MainWindow  │  │ PreviewWidget│  │ Edit Form │ │
│  └──────┬───────┘  └──────┬───────┘  └─────┬─────┘ │
└─────────┼─────────────────┼─────────────────┼──────┘
          │                 │                 │
┌─────────▼─────────────────▼─────────────────▼──────┐
│                  Application Layer                   │
│              AppController (Orchestrator)             │
│                                                      │
│  Responsibilities:                                   │
│  • Coordinate the full document processing pipeline  │
│  • Manage UI state and user interactions             │
│  • Handle error reporting and validation feedback    │
└──────────────────────────┬───────────────────────────┘
                           │
┌──────────────────────────▼───────────────────────────┐
│                   Service Layer                       │
│                                                      │
│  ┌─────────┐  ┌──────────┐  ┌───────────┐           │
│  │ Reader  │─▶│ Parsers  │─▶│ Validator │           │
│  │         │  │          │  │           │           │
│  │ • .docx │  │ • Header │  │ • Check   │           │
│  │ • .doc* │  │ • Parties│  │   fields  │           │
│  │         │  │ • Object │  │ • Types   │           │
│  │         │  │ • Period │  │ • Totals  │           │
│  │         │  │ • Tables │  │           │           │
│  │         │  │ • Totals │  │           │           │
│  └─────────┘  └──────────┘  └─────┬─────┘           │
│                                    │                 │
│  ┌─────────────────────────────────▼───────────────┐│
│  │              Generators                          ││
│  │  ┌────────┐  ┌───────┐  ┌───────┐               ││
│  │  │  Act   │  │ KS-2  │  │ KS-3  │               ││
│  │  │ .docx  │  │ .xlsx │  │ .xlsx │               ││
│  │  └────────┘  └───────┘  └───────┘               ││
│  └──────────────────────────────────────────────────┘│
└──────────────────────────────────────────────────────┘
```

## Core Pipeline

### 1. Document Reading

- Parses `.docx` files using `python-docx`
- Extracts paragraphs and tables in document order
- Builds a structured `ContractDocument` model

### 2. Classification

- Analyzes document structure to determine:
  - **Party types**: Organization vs Individual Entrepreneur
  - **Table structure**: Flat vs Sectional estimate layout
- Classification drives parser selection and template choice

### 3. Data Extraction (Parsing Pipeline)

Each parser focuses on a specific data domain:

| Parser | Extracts |
|--------|----------|
| HeaderParser | Contract number, date, city |
| PartiesParser | Customer and executor details (name, INN, OGRN, bank, address) |
| ObjectParser | Construction object name, address, inventory number |
| PeriodParser | Contract dates, work period |
| TableParser | Estimate line items (name, quantity, price, totals) |
| TotalsParser | Summary amounts, VAT calculations |

### 4. Validation

- Checks all required fields are present
- Validates data consistency (e.g., ORG needs KPP, IP needs OGRNIP)
- Returns status: OK / WARNING / ERROR

### 5. Document Generation

- **Act** — Word document with contract details, parties, estimate summary, signatures
- **KS-2** — Excel form with full estimate table, totals, payment section
- **KS-3** — Excel form with work summary by category

All generators use template files:
- Placeholders like `{{contract_number}}` are replaced
- `{{ROW_TEMPLATE}}` rows are cloned for each estimate item
- Formatting is preserved from the original template

## Design Patterns

| Pattern | Usage |
|---------|-------|
| **Strategy** | Parsers and generators are swappable via interfaces |
| **Dependency Injection** | All services injected into AppController |
| **Template Method** | Each parser follows extract → validate → return |
| **Factory** | TemplateLoader selects correct template by contract type |

## Demo vs Production

This demo version uses **stub implementations** that return sample data. The production version includes:

- **20+ regex patterns** per parser for different contract formats
- **Table structure analysis** with section header detection
- **Party detection heuristics** for ORG vs IP identification
- **Multi-format date parsing** with Russian locale support
- **Bank details extraction** with validation
- **Address normalization** and shortening

The interface contracts remain the same — swap stubs for production implementations to get full functionality.

## Technology Choices

| Choice | Reason |
|--------|--------|
| PySide6 | Rich desktop UI, native Windows integration |
| python-docx | Direct .docx manipulation without Word dependency |
| openpyxl | Full Excel file control (no Excel required) |
| pydantic | Data validation and serialization |
| pywin32 | Word automation for PDF preview generation |
| pytest | Comprehensive test coverage |
