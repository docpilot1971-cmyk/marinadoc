from __future__ import annotations

from datetime import date
from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field

from app.models.enums import PartyType, RowGroupingMode, RowType, ValidationStatus


class AppBaseModel(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class DocumentData(AppBaseModel):
    contract_number: str | None = None
    contract_date: date | None = None
    document_city: str | None = None
    act_date: date | None = None


class PartyData(AppBaseModel):
    type: PartyType | None = None
    full_name: str | None = None
    short_name: str | None = None
    representative_name: str | None = None
    representative_position: str | None = None
    representative_basis: str | None = None
    inn: str | None = None
    kpp: str | None = None
    ogrn: str | None = None
    ogrnip: str | None = None
    address: str | None = None
    bank_name: str | None = None
    rs: str | None = None
    ks: str | None = None
    bik: str | None = None
    passport: str | None = None
    registration: str | None = None
    tax_office: str | None = None


class ObjectData(AppBaseModel):
    object_name: str | None = None
    object_address: str | None = None
    object_inventory_no: str | None = None
    object_cadastral_no: str | None = None


class PeriodData(AppBaseModel):
    work_start_date: date | None = None
    work_end_date_plan: date | None = None
    work_end_date_fact: date | None = None
    reporting_period: str | None = None


class EstimateRow(AppBaseModel):
    row_grouping_mode: RowGroupingMode = RowGroupingMode.FLAT
    row_type: RowType = RowType.ITEM
    row_section: str | None = None
    row_number: str | None = None
    row_name: str | None = None
    row_unit: str | None = None
    row_quantity: Decimal | None = None
    row_price: Decimal | None = None
    row_amount: Decimal | None = None
    row_completion_date: date | None = None
    row_sort_index: int = 0


class TotalsData(AppBaseModel):
    works_total: Decimal = Decimal("0")
    materials_total: Decimal = Decimal("0")
    transport_total: Decimal = Decimal("0")
    travel_total: Decimal = Decimal("0")
    total_without_vat: Decimal = Decimal("0")
    vat_rate: Decimal = Decimal("0")
    vat_amount: Decimal = Decimal("0")
    total_with_vat: Decimal = Decimal("0")
    total_in_words: str | None = None


class ExtractionResult(AppBaseModel):
    document: DocumentData = Field(default_factory=DocumentData)
    customer: PartyData = Field(default_factory=PartyData)
    executor: PartyData = Field(default_factory=PartyData)
    object_data: ObjectData = Field(default_factory=ObjectData)
    period: PeriodData = Field(default_factory=PeriodData)
    rows: list[EstimateRow] = Field(default_factory=list)
    totals: TotalsData = Field(default_factory=TotalsData)
    validation_status: ValidationStatus = ValidationStatus.WARNING
    validation_messages: list[str] = Field(default_factory=list)
