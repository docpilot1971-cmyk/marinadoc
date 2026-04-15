from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from app.services.contract_document import ContractDocument


class IContractReader(ABC):
    @abstractmethod
    def read(self, file_path: Path) -> ContractDocument:
        """Read contract from source file and return structured representation."""
