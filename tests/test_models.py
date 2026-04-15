from datetime import date
from decimal import Decimal

from app.models import (
    DocumentData,
    EstimateRow,
    ExtractionResult,
    PartyData,
    PartyType,
    RowGroupingMode,
    RowType,
    TotalsData,
    ValidationStatus,
)


def test_document_data_model() -> None:
    data = DocumentData(
        contract_number="42-2026",
        contract_date=date(2026, 3, 1),
        document_city="Moscow",
        act_date=date(2026, 3, 31),
    )
    assert data.contract_number == "42-2026"
    assert data.contract_date == date(2026, 3, 1)


def test_party_type_enum() -> None:
    party = PartyData(type=PartyType.ORG, full_name='OOO "Test"')
    assert party.type == PartyType.ORG


def test_estimate_row_defaults() -> None:
    row = EstimateRow()
    assert row.row_grouping_mode == RowGroupingMode.FLAT
    assert row.row_type == RowType.ITEM
    assert row.row_sort_index == 0


def test_totals_model_decimal_values() -> None:
    totals = TotalsData(total_with_vat=Decimal("123.45"))
    assert totals.total_with_vat == Decimal("123.45")


def test_extraction_result_defaults() -> None:
    result = ExtractionResult()
    assert result.validation_status == ValidationStatus.WARNING
    assert result.rows == []
