from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from app.models import ExtractionResult


class IActWordGenerator(ABC):
    @abstractmethod
    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        """Generate Word act document."""


class IKS2ExcelGenerator(ABC):
    @abstractmethod
    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        """Generate KS-2 Excel document."""


class IKS3ExcelGenerator(ABC):
    @abstractmethod
    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        """Generate KS-3 Excel document."""
