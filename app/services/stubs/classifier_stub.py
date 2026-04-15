"""
Contract classifier stub — simplified for public demo.

This module demonstrates the classification pipeline architecture.
The production version uses multi-strategy detection for party types
(ORG/IP), table grouping modes, and contract structure analysis.

Full version available on request.
"""
from __future__ import annotations

import logging

from app.models import PartyType, RowGroupingMode
from app.services.classification import ContractClassification
from app.services.contract_document import ContractDocument
from app.services.interfaces import IContractTypeClassifier

logger = logging.getLogger(__name__)


class ContractTypeClassifierStub(IContractTypeClassifier):
    """Classify contract type and table structure.

    Production version uses section-based text analysis, table structure
    detection, and multi-pattern party type detection.
    Simplified here for demo.
    """

    def classify(self, contract: ContractDocument) -> ContractClassification:
        # simplified for public demo
        # In production: analyses customer/executor sections, detects IP vs ORG
        # by multiple indicators (ОГРНИП, «ИП», registration patterns)
        classification = ContractClassification(
            customer_type=PartyType.ORG,
            executor_type=PartyType.ORG,
            table_grouping_mode=RowGroupingMode.FLAT,
        )
        logger.info("Contract classified (demo): %s", classification)
        return classification
