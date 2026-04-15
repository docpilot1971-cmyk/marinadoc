from __future__ import annotations

from abc import ABC, abstractmethod

from app.services.classification import ContractClassification
from app.services.contract_document import ContractDocument


class IContractTypeClassifier(ABC):
    @abstractmethod
    def classify(self, contract: ContractDocument) -> ContractClassification:
        """Return classification data for parties and estimate table structure."""
