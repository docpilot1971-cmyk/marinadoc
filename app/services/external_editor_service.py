from __future__ import annotations

import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)


class ExternalEditorService:
    def open_in_word(self, path: Path) -> None:
        if not path.exists():
            raise FileNotFoundError(path)
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]

            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True
            word.Documents.Open(str(path.resolve()))
            logger.info("Opened in Word: %s", path)
            # NOTE: Не вызываем CoUninitialize, так как Word остаётся открытым
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to open Word file: %s", path)
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
                return
            except Exception as fallback_exc:  # noqa: BLE001
                raise RuntimeError(f"Cannot open Word file: {path}") from fallback_exc

    def open_in_excel(self, path: Path) -> None:
        if not path.exists():
            raise FileNotFoundError(path)
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]

            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            excel.Workbooks.Open(str(path.resolve()))
            logger.info("Opened in Excel: %s", path)
            # NOTE: Не вызываем CoUninitialize, так как Excel остаётся открытым
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to open Excel file: %s", path)
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
                return
            except Exception as fallback_exc:  # noqa: BLE001
                raise RuntimeError(f"Cannot open Excel file: {path}") from fallback_exc

    def open_for_edit(self, path: Path) -> None:
        suffix = path.suffix.lower()
        if suffix == ".docx":
            self.open_in_word(path)
            return
        if suffix == ".xlsx":
            self.open_in_excel(path)
            return
        raise ValueError(f"Unsupported file format for external editing: {suffix}")
