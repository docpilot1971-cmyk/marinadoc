"""
Parser stubs — simplified for public demo.

This module demonstrates the document parsing pipeline architecture.
The full production version contains advanced heuristics, multi-pattern
extraction rules, and domain-specific parsing logic for construction
contracts.

Full version available on request.
"""
from __future__ import annotations

import logging
import re
from datetime import date

from app.models import (
    DocumentData,
    EstimateRow,
    ObjectData,
    PartyData,
    PeriodData,
    RowGroupingMode,
    RowType,
    TotalsData,
)
from app.services.contract_document import ContractDocument
from app.services.interfaces import (
    IHeaderParser,
    IObjectParser,
    IPartiesParser,
    IPeriodParser,
    ITableParser,
    ITotalsParser,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Header Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class HeaderParserStub(IHeaderParser):
    """Extract basic document header: contract number, date, city.

    Production version uses multiple regex strategies, fallback chains,
    and context-aware extraction. Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> DocumentData:
        # simplified for public demo
        header_text = "\n".join(contract.paragraphs[:20])

        contract_number = self._extract_contract_number(header_text)
        contract_date = self._extract_contract_date(header_text)
        city = self._extract_city(header_text)

        return DocumentData(
            contract_number=contract_number,
            contract_date=contract_date,
            document_city=city,
            act_date=date.today(),
        )

    @staticmethod
    def _extract_contract_number(text: str) -> str | None:
        match = re.search(r"№\s*([A-Za-zА-Яа-яЁё0-9\-/]+)", text)
        return match.group(1).strip() if match else None

    @staticmethod
    def _extract_contract_date(text: str) -> date | None:
        match = re.search(r"(\d{2})[.\-/](\d{2})[.\-/](\d{4})", text)
        if match:
            try:
                return date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
            except (ValueError, OverflowError):
                return None
        return None

    @staticmethod
    def _extract_city(text: str) -> str | None:
        match = re.search(r"г\.\s*([А-ЯЁ][А-Яа-яЁё\-]{2,40})", text)
        return match.group(1).strip() if match else None


# ---------------------------------------------------------------------------
# Parties Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class PartiesParserStub(IPartiesParser):
    """Extract customer and executor party data.

    Production version includes: table-based requisites extraction,
    multi-pattern name extraction (ORG/IP), representative detection,
    bank details parsing, address normalisation, and basis detection.
    Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> tuple[PartyData, PartyData]:
        # simplified for public demo
        from app.models.enums import PartyType

        customer = PartyData(
            type=PartyType.ORG,
            full_name='ООО "Заказчик"',
            short_name="ООО Заказчик",
            representative_name="Иванов Иван Иванович",
            representative_position="Генеральный директор",
            representative_basis="Устава",
            inn="7701234567",
            kpp="770101001",
            ogrn="1177746000000",
            ogrnip=None,
            address="г. Москва, ул. Примерная, д. 1",
            bank_name="ПАО Сбербанк",
            rs="40702810000000000001",
            ks="30101810400000000225",
            bik="044525225",
            passport=None,
            registration=None,
            tax_office=None,
        )
        executor = PartyData(
            type=PartyType.ORG,
            full_name='ООО "Подрядчик"',
            short_name="ООО Подрядчик",
            representative_name="Петров Пётр Петрович",
            representative_position="Генеральный директор",
            representative_basis="Устава",
            inn="7709876543",
            kpp="770901001",
            ogrn="1187746000000",
            ogrnip=None,
            address="г. Москва, ул. Строителей, д. 10",
            bank_name="ПАО Сбербанк",
            rs="40702810000000000002",
            ks="30101810400000000225",
            bik="044525225",
            passport=None,
            registration=None,
            tax_office=None,
        )
        return customer, executor


# ---------------------------------------------------------------------------
# Object Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class ObjectParserStub(IObjectParser):
    """Extract construction object information.

    Production version uses multiple extraction strategies for object
    name, address, inventory and cadastral numbers.
    Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> ObjectData:
        # simplified for public demo
        return ObjectData(
            object_name="Объект (демо)",
            object_address="г. Москва, демо-адрес",
            object_inventory_no="ИНВ-001",
            object_cadastral_no=None,
        )


# ---------------------------------------------------------------------------
# Period Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class PeriodParserStub(IPeriodParser):
    """Extract contract period: start/end dates, work period.

    Production version handles multiple date formats, relative date
    expressions, and table-based period extraction.
    Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> PeriodData:
        # simplified for public demo
        return PeriodData(
            contract_start=date(2024, 1, 1),
            contract_end=date(2024, 12, 31),
            work_start=date(2024, 1, 15),
            work_end=date(2024, 6, 30),
        )


# ---------------------------------------------------------------------------
# Table Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class TableParserStub(ITableParser):
    """Extract estimate rows from contract tables.

    Production version handles sectional and flat table layouts,
    section header detection, service grouping, numbering extraction,
    unit parsing, and multi-row estimate entries.
    Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> list[EstimateRow]:
        # simplified for public demo
        return [
            EstimateRow(
                num="1",
                name="Демонстрационная работа 1",
                unit="шт",
                quantity=10.0,
                price=1500.00,
                total=15000.00,
                row_type=RowType.WORK,
                grouping_mode=RowGroupingMode.FLAT,
                section_name=None,
                section_num=None,
            ),
            EstimateRow(
                num="2",
                name="Демонстрационная работа 2",
                unit="м2",
                quantity=50.0,
                price=800.00,
                total=40000.00,
                row_type=RowType.WORK,
                grouping_mode=RowGroupingMode.FLAT,
                section_name=None,
                section_num=None,
            ),
            EstimateRow(
                num="3",
                name="Демонстрационный материал",
                unit="компл",
                quantity=1.0,
                price=25000.00,
                total=25000.00,
                row_type=RowType.MATERIAL,
                grouping_mode=RowGroupingMode.FLAT,
                section_name=None,
                section_num=None,
            ),
        ]


# ---------------------------------------------------------------------------
# Totals Parser (simplified for public demo)
# ---------------------------------------------------------------------------

class TotalsParserStub(ITotalsParser):
    """Extract contract totals: sum, VAT, breakdown.

    Production version handles multiple table layouts, VAT calculations,
    and summary row detection.
    Simplified here for demo.
    """

    def parse(self, contract: ContractDocument) -> TotalsData:
        # simplified for public demo
        return TotalsData(
            total_with_vat=80000.00,
            total_without_vat=None,
            vat_amount=None,
            vat_rate=None,
            works_total=55000.00,
            materials_total=25000.00,
        )
