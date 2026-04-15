from enum import StrEnum


class PartyType(StrEnum):
    ORG = "ORG"
    IP = "IP"


class RowGroupingMode(StrEnum):
    SECTIONAL = "SECTIONAL"
    FLAT = "FLAT"


class RowType(StrEnum):
    SECTION = "SECTION"
    ITEM = "ITEM"
    SUBTOTAL = "SUBTOTAL"
    COST_ONLY = "COST_ONLY"
    GRAND_TOTAL = "GRAND_TOTAL"


class ValidationStatus(StrEnum):
    OK = "OK"
    WARNING = "WARNING"
    ERROR = "ERROR"
