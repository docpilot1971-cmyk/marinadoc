"""Test BELGOR contract parsing."""
from pathlib import Path
import sys
sys.path.append('.')

from app.services.stubs.reader_stub import ContractReaderStub
from app.services.stubs.parsers_stub import PartiesParserStub
from app.services.stubs.classifier_stub import ContractTypeClassifierStub

contract_path = Path('templates/incoming/Договор подряда с БЕЛОГОР (7).docx')
print(f"Testing: {contract_path}")

reader = ContractReaderStub()
contract = reader.read(contract_path)
cls = ContractTypeClassifierStub().classify(contract)
parties_parser = PartiesParserStub()

customer, executor = parties_parser.parse(contract)

print(f"\nExecutor (Contractor): {executor.full_name}")
print(f"  Registration: {executor.registration}")
print(f"  IFNS: {executor.tax_office}")
print(f"  Passport: {executor.passport}")
