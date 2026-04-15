from __future__ import annotations

from dataclasses import dataclass

from app.models import PartyType, RowGroupingMode


@dataclass(slots=True)
class ContractClassification:
    customer_type: PartyType
    executor_type: PartyType
    table_grouping_mode: RowGroupingMode
