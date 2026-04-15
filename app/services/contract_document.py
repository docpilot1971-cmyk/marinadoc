from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path


class BlockKind(StrEnum):
    PARAGRAPH = "PARAGRAPH"
    TABLE = "TABLE"


@dataclass(slots=True)
class ParagraphBlock:
    kind: BlockKind = BlockKind.PARAGRAPH
    text: str = ""


@dataclass(slots=True)
class TableBlock:
    kind: BlockKind = BlockKind.TABLE
    rows: list[list[str]] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(" | ".join(c for c in row if c) for row in self.rows if any(row))


@dataclass(slots=True)
class ContractDocument:
    file_path: Path
    paragraphs: list[str] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)
    blocks: list[ParagraphBlock | TableBlock] = field(default_factory=list)

    @property
    def full_text(self) -> str:
        lines: list[str] = []
        for block in self.blocks:
            if isinstance(block, ParagraphBlock):
                if block.text:
                    lines.append(block.text)
            else:
                block_text = block.text
                if block_text:
                    lines.append(block_text)
        return "\n".join(lines)
