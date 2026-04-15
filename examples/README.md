# Examples

This directory contains sample input and output files demonstrating the MarinaDoc workflow.

## sample_input/

Sample contract files you can use to test the application:

- `sample_contract.docx` — A demo construction contract with standard ORG-ORG layout

## sample_output/

Example documents generated from the sample input:

- `Акт_001_15.04.2024.docx` — Generated Act
- `КС2_001_15.04.2024.xlsx` — Generated KS-2 form
- `КС3_001_15.04.2024.xlsx` — Generated KS-3 form

## Creating Your Own Test Documents

To create a test contract:

1. Use any `.docx` file with:
   - Contract header (number, date, city)
   - Party definitions (Заказчик / Подрядчик sections)
   - Estimate table (works, materials, quantities, prices)
   - Party requisites (INN, bank details, addresses)

2. Place it in `sample_input/` and open in MarinaDoc

> **Note:** This demo version returns sample data regardless of the input file. The full production version extracts actual data from the contract.
