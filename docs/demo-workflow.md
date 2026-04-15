# Demo Workflow Guide

## What This Demo Shows

This demo version illustrates the **complete document automation pipeline** from contract upload to document generation. While the parsing returns sample data (not real extraction), the full workflow is functional.

---

## Complete Workflow

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│   1. Open    │────▶│   2. Parse   │────▶│   3. Edit    │────▶│   4. Review  │
│   Contract   │     │   & Extract  │     │   & Validate │     │   & Confirm  │
│   (.docx)    │     │   Data       │     │   Data       │     │   Generation │
└──────────────┘     └──────────────┘     └──────────────┘     └──────┬───────┘
                                                                      │
                                                       ┌──────────────▼───────┐
                                                       │   5. Generate Docs   │
                                                       │   • Act (.docx)      │
                                                       │   • KS-2 (.xlsx)     │
                                                       │   • KS-3 (.xlsx)     │
                                                       └──────────────────────┘
```

---

## Step 1: Open a Contract

1. Launch the app: `python main.py`
2. Click **Open Contract** or use `File → Open`
3. Select a `.docx` contract file

**What happens:**
- The document is read and parsed into paragraphs and tables
- The contract is classified (ORG/IP type)
- In demo mode: sample data is loaded

---

## Step 2: Data Extraction

**What happens behind the scenes:**

| Stage | Parser | Extracts |
|-------|--------|----------|
| Header | HeaderParser | Contract number, date, city |
| Parties | PartiesParser | Customer & executor details |
| Object | ObjectParser | Construction object info |
| Period | PeriodParser | Contract and work dates |
| Tables | TableParser | Estimate line items |
| Totals | TotalsParser | Summary amounts |

**Demo vs Production:**
- **Demo**: Returns pre-filled sample data
- **Production**: Extracts actual data from the contract using advanced heuristics

---

## Step 3: Edit & Validate

The editing form shows all extracted data:
- **Document section**: Contract number, date, city
- **Customer section**: Name, INN, KPP, address, bank details
- **Executor section**: Name, INN, KPP, address, bank details
- **Object section**: Object name, address
- **Estimate table**: Line items with quantities and prices
- **Totals**: Summary amounts

**Validation checks:**
- All required fields present
- ORG type → KPP required
- IP type → OGRNIP required
- Estimate table has rows
- Totals are positive

---

## Step 4: Review & Confirm

Review the extracted and edited data. You can:
- Modify any field
- Add/remove estimate rows
- Fix validation errors (shown in the UI)

---

## Step 5: Generate Documents

Click **Generate** to create the full document set:

### Generated Files

| File | Format | Content |
|------|--------|---------|
| `Акт_{number}.docx` | Word | Act with parties, estimate summary, signatures |
| `КС2_{number}.xlsx` | Excel | KS-2 form with full estimate table |
| `КС3_{number}.xlsx` | Excel | KS-3 form with work summary |

### File Naming

Files are named using the pattern:
- `Акт_{contract_number}_{date}.docx`
- `КС2_{contract_number}_{date}.xlsx`
- `КС3_{contract_number}_{date}.xlsx`

---

## What's Simplified in This Demo

| Feature | Demo | Production |
|---------|------|------------|
| Contract reading | ✅ Full | Full |
| Document preview | ✅ Full | Full |
| UI editing | ✅ Full | Full |
| Data extraction | Sample data | Real extraction |
| Party detection | Static | Dynamic ORG/IP |
| Table parsing | Sample rows | Full table analysis |
| Validation | ✅ Full | Full |
| Document generation | ✅ Full | Full |
| Template support | Demo templates | All template types |

---

## Production Capabilities (Not in Demo)

The production version includes:

1. **Multi-format contract parsing**
   - 20+ regex patterns per data field
   - Support for various contract layouts
   - Table-based and text-based requisites extraction

2. **Smart classification**
   - Automatic ORG/IP detection
   - Sectional vs flat table mode detection
   - Handles mixed contract formats

3. **Advanced estimate parsing**
   - Section header detection
   - Service grouping by category
   - Handles nested tables and merged cells

4. **Complete party extraction**
   - INN, KPP, OGRN/OGRNIP extraction
   - Bank details (account, BIK, correspondent account)
   - Address normalization
   - Representative and basis detection

5. **Production templates**
   - ORG-ORG and ORG-IP template variants
   - Proper formatting preservation
   - Full KS-2/KS-3 form compliance

---

## Getting the Full Version

This demo demonstrates the architecture and workflow. For the full production version with advanced parsing capabilities, contact the author.
