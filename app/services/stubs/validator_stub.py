from __future__ import annotations

import logging

from app.models import ExtractionResult, ValidationStatus
from app.services.interfaces import IExtractionValidator

logger = logging.getLogger(__name__)


class ExtractionValidatorStub(IExtractionValidator):
    def validate(self, data: ExtractionResult) -> ExtractionResult:
        messages: list[str] = []

        if not data.document.contract_number:
            messages.append("Contract number is missing.")
        if not data.document.contract_date:
            messages.append("Contract date is missing.")
        if not data.document.act_date:
            messages.append("Act date is missing.")

        if not data.customer.full_name:
            messages.append("Customer full name is missing.")
        if not data.customer.inn:
            messages.append("Customer INN is missing.")

        if not data.executor.full_name:
            messages.append("Executor full name is missing.")
        if not data.executor.inn:
            messages.append("Executor INN is missing.")

        if data.customer.type and data.customer.type.value == "ORG" and not data.customer.kpp:
            messages.append("Customer KPP is missing for ORG.")
        if data.executor.type and data.executor.type.value == "ORG" and not data.executor.kpp:
            messages.append("Executor KPP is missing for ORG.")

        if data.customer.type and data.customer.type.value == "IP":
            data.customer.kpp = None
            if not data.customer.ogrnip:
                messages.append("Customer OGRNIP is missing for IP.")
        if data.executor.type and data.executor.type.value == "IP":
            data.executor.kpp = None
            if not data.executor.ogrnip:
                messages.append("Executor OGRNIP is missing for IP.")

        if not data.rows:
            messages.append("Estimate table rows were not detected.")

        if data.totals.total_without_vat <= 0:
            messages.append("Total without VAT is missing or invalid.")
        if data.totals.vat_amount < 0:
            messages.append("VAT amount is invalid.")
        if data.totals.total_with_vat <= 0:
            messages.append("Total with VAT is missing or invalid.")

        has_error = not data.rows or data.totals.total_with_vat <= 0
        data.validation_messages = messages
        if has_error:
            data.validation_status = ValidationStatus.ERROR
        elif messages:
            data.validation_status = ValidationStatus.WARNING
        else:
            data.validation_status = ValidationStatus.OK

        logger.info("Validation stub completed with %s message(s).", len(messages))
        return data
