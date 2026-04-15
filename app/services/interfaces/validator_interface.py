from __future__ import annotations

from abc import ABC, abstractmethod

from app.models import ExtractionResult


class IExtractionValidator(ABC):
    @abstractmethod
    def validate(self, data: ExtractionResult) -> ExtractionResult:
        """Validate extraction result and update validation status/messages."""
