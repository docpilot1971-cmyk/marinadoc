from datetime import date
from decimal import Decimal
from pathlib import Path

from app.models import PartyType, RowGroupingMode
from app.services.contract_document import ContractDocument, ParagraphBlock, TableBlock
from app.services.stubs.classifier_stub import ContractTypeClassifierStub
from app.services.stubs.parsers_stub import (
    HeaderParserStub,
    PartiesParserStub,
    TableParserStub,
    TotalsParserStub,
)


def _build_contract() -> ContractDocument:
    paragraphs = [
        'Договор подряда № 45-2025 от 12.03.2025 г. Москва',
        'Заказчик: ООО "Ромашка", в лице Иванова И.И., действующего на основании Устава',
        "Подрядчик: Индивидуальный предприниматель Петров П.П., действующий на основании свидетельства",
        "Реквизиты",
        "Заказчик",
        "ИНН 7701234567 КПП 770101001 р/с 40702810100000000001 к/с 30101810400000000225 БИК 044525225",
        "Подрядчик",
        "ИНН 500123456789 ОГРНИП 320500000000123 р/с 40802810900000000002 к/с 30101810100000000666 БИК 044525666",
    ]
    table = [
        ["№", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма"],
        ["", "Работы", "", "", "", ""],
        ["1", "Монтаж", "шт", "2", "1000", "2000"],
        ["", "Материалы", "", "", "", ""],
        ["2", "Кабель", "м", "10", "50", "500"],
        ["", "Итого без НДС", "", "", "", "2500"],
        ["", "НДС 20%", "", "", "", "500"],
        ["", "Итого с НДС", "", "", "", "3000"],
    ]
    blocks = [
        ParagraphBlock(text=paragraphs[0]),
        ParagraphBlock(text=paragraphs[1]),
        ParagraphBlock(text=paragraphs[2]),
        ParagraphBlock(text=paragraphs[3]),
        ParagraphBlock(text=paragraphs[4]),
        ParagraphBlock(text=paragraphs[5]),
        ParagraphBlock(text=paragraphs[6]),
        ParagraphBlock(text=paragraphs[7]),
        TableBlock(rows=table),
    ]
    return ContractDocument(
        file_path=Path("sample.docx"),
        paragraphs=paragraphs,
        tables=[table],
        blocks=blocks,
    )


def test_header_parser_extracts_number_and_date() -> None:
    contract = _build_contract()
    parsed = HeaderParserStub().parse(contract)
    assert parsed.contract_number == "45-2025"
    assert parsed.contract_date == date(2025, 3, 12)


def test_classifier_detects_ip_and_sectional() -> None:
    contract = _build_contract()
    cls = ContractTypeClassifierStub().classify(contract)
    assert cls.customer_type == PartyType.ORG
    assert cls.executor_type == PartyType.IP
    assert cls.table_grouping_mode == RowGroupingMode.SECTIONAL


def test_parties_parser_extracts_requisites() -> None:
    contract = _build_contract()
    customer, executor = PartiesParserStub().parse(contract)
    assert customer.inn == "7701234567"
    assert customer.kpp == "770101001"
    assert executor.inn == "500123456789"
    assert executor.ogrnip == "320500000000123"


def test_parties_parser_extracts_customer_name_and_representative() -> None:
    contract = _build_contract()
    customer, executor = PartiesParserStub().parse(contract)

    assert customer.full_name == 'ООО "Ромашка"'
    assert customer.representative_name == "Иванова И.И."
    assert customer.representative_basis == "Устава"


def test_parties_parser_extracts_ip_executor_correctly() -> None:
    contract = _build_contract()
    customer, executor = PartiesParserStub().parse(contract)

    assert executor.full_name == "Индивидуальный предприниматель Петров П.П."
    assert executor.representative_name == "Петров П.П."
    assert executor.ogrnip == "320500000000123"
    assert executor.kpp is None


def test_parties_parser_does_not_replace_customer_with_representative() -> None:
    contract = _build_contract()
    customer, executor = PartiesParserStub().parse(contract)

    assert customer.full_name != customer.representative_name


def test_table_parser_extracts_rows() -> None:
    contract = _build_contract()
    rows = TableParserStub().parse(contract, RowGroupingMode.SECTIONAL)
    item_rows = [r for r in rows if r.row_type.value == "ITEM"]
    assert len(item_rows) >= 2
    assert item_rows[0].row_name == "Монтаж"


def test_totals_parser_extracts_totals() -> None:
    contract = _build_contract()
    rows = TableParserStub().parse(contract, RowGroupingMode.SECTIONAL)
    totals = TotalsParserStub().parse(contract, rows)
    assert totals.total_without_vat == Decimal("2500")
    assert totals.vat_rate == Decimal("20")
    assert totals.vat_amount == Decimal("500")
    assert totals.total_with_vat == Decimal("3000")


def test_totals_parser_defaults_to_no_vat_when_not_explicit() -> None:
    paragraphs = [
        "Договор подряда № 99 от 01.02.2026",
        'Заказчик: ООО "Ромашка"',
        "Подрядчик: ИП Петров П.П.",
    ]
    table = [
        ["№", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма"],
        ["1", "Монтаж", "шт", "2", "1000", "2000"],
        ["2", "Кабель", "м", "10", "50", "500"],
        ["", "Итого", "", "", "", "2500"],
    ]
    contract = ContractDocument(
        file_path=Path("sample_no_vat.docx"),
        paragraphs=paragraphs,
        tables=[table],
        blocks=[TableBlock(rows=table)],
    )

    rows = TableParserStub().parse(contract, RowGroupingMode.FLAT)
    totals = TotalsParserStub().parse(contract, rows)

    assert totals.total_without_vat == Decimal("2500.00")
    assert totals.vat_rate == Decimal("0")
    assert totals.vat_amount == Decimal("0")
    assert totals.total_with_vat == Decimal("2500.00")
