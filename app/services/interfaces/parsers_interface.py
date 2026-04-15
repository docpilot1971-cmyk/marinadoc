from __future__ import annotations

from abc import ABC, abstractmethod

from app.models import DocumentData, EstimateRow, ObjectData, PartyData, PeriodData, TotalsData
from app.models.enums import RowGroupingMode
from app.services.contract_document import ContractDocument


class IHeaderParser(ABC):
    @abstractmethod
    def parse(self, contract: ContractDocument) -> DocumentData:
        """Parse contract header block."""


class IPartiesParser(ABC):
    @abstractmethod
    def parse(self, contract: ContractDocument) -> tuple[PartyData, PartyData]:
        """Parse customer and executor blocks."""


class IObjectParser(ABC):
    @abstractmethod
    def parse(self, contract: ContractDocument) -> ObjectData:
        """Parse object details block."""


class IPeriodParser(ABC):
    @abstractmethod
    def parse(self, contract: ContractDocument) -> PeriodData:
        """Parse work/reporting period details."""


class ITableParser(ABC):
    @abstractmethod
    def parse(
        self,
        contract: ContractDocument,
        grouping_hint: RowGroupingMode | None = None,
    ) -> list[EstimateRow]:
        """Parse estimate table rows."""


class ITotalsParser(ABC):
    @abstractmethod
    def parse(self, contract: ContractDocument, rows: list[EstimateRow]) -> TotalsData:
        """Parse totals or derive them from rows."""
