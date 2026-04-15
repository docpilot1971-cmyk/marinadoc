from __future__ import annotations

import logging
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


class DocumentPreviewService:
    def __init__(self) -> None:
        self._temp_files: set[Path] = set()

    def preview_document(self, input_path: Path) -> Path:
        suffix = input_path.suffix.lower()
        if suffix in {".docx", ".doc"}:
            return self.convert_docx_to_pdf(input_path)
        if suffix == ".pdf":
            return input_path
        raise ValueError(f"Unsupported preview source format: {suffix}")

    def convert_docx_to_pdf(self, input_path: Path) -> Path:
        if not input_path.exists():
            raise FileNotFoundError(input_path)

        temp_dir = Path(tempfile.mkdtemp(prefix="contract_preview_"))
        output_pdf = temp_dir / f"{input_path.stem}.pdf"

        word = None
        document = None
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
            import time

            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(str(input_path.resolve()), ReadOnly=True)
            
            # Небольшая задержка для полной загрузки документа
            time.sleep(0.2)
            
            document.ExportAsFixedFormat(str(output_pdf.resolve()), 17)
            logger.info("Contract preview PDF created: %s", output_pdf)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to convert contract to PDF: %s", input_path)
            raise RuntimeError(f"Cannot convert DOCX to PDF: {input_path}") from exc
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
            if word is not None:
                try:
                    word.Quit(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
            
            # Освобождаем COM-объекты
            import gc
            document = None
            word = None
            gc.collect()
            
            try:
                import pythoncom  # type: ignore[import-not-found]
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass

        if not output_pdf.exists():
            raise RuntimeError(f"PDF preview file was not created: {output_pdf}")

        self._temp_files.add(output_pdf)
        return output_pdf

    def cleanup_temp_preview(self, file_path: Path) -> None:
        import time
        try:
            if file_path.exists():
                parent = file_path.parent
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    # Файл заблокирован, попробуем позже
                    time.sleep(1)
                    try:
                        file_path.unlink(missing_ok=True)
                    except Exception:
                        pass
                if parent.exists() and not any(parent.iterdir()):
                    try:
                        parent.rmdir()
                    except Exception:
                        pass
        except Exception:  # noqa: BLE001
            logger.exception("Failed to cleanup temp preview: %s", file_path)
        finally:
            self._temp_files.discard(file_path)

    def cleanup_all(self) -> None:
        for file_path in list(self._temp_files):
            self.cleanup_temp_preview(file_path)
        # safety cleanup for empty temp dirs left by failures
        tmp_root = Path(tempfile.gettempdir())
        for child in tmp_root.glob("contract_preview_*"):
            if child.is_dir():
                try:
                    if not any(child.iterdir()):
                        child.rmdir()
                except Exception:  # noqa: BLE001
                    pass
