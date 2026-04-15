from __future__ import annotations

import logging
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


class OutputPreviewService:
    def __init__(self, preview_dir: Path | None = None) -> None:
        self.preview_dir = preview_dir or Path(tempfile.gettempdir()) / "contract_output_preview"
        try:
            self.preview_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            logger.warning("Failed to create preview directory %s: %s", self.preview_dir, exc)
        self._temp_pdfs: set[Path] = set()

    def build_preview(self, input_path: Path) -> Path:
        suffix = input_path.suffix.lower()
        if suffix == ".docx":
            return self.convert_docx_to_pdf(input_path)
        if suffix == ".xlsx":
            return self.convert_xlsx_to_pdf(input_path)
        if suffix == ".pdf":
            return input_path
        raise ValueError(f"Unsupported output format for preview: {suffix}")

    def convert_docx_to_pdf(self, input_path: Path) -> Path:
        import uuid
        output_pdf = self.preview_dir / f"{uuid.uuid4().hex[:8]}_{input_path.stem}.pdf"
        word = None
        doc = None
        
        import shutil
        import time
        import uuid

        # Retry loop to handle transient file locks
        for attempt in range(3):
            temp_dir = Path(tempfile.mkdtemp(prefix="contract_preview_"))
            # Use a unique filename to avoid collisions with locked files from previous failed runs
            unique_name = f"{uuid.uuid4().hex[:8]}_{input_path.stem}{input_path.suffix}"
            temp_copy = temp_dir / unique_name
            
            try:
                # Copy file to temp location
                shutil.copy2(str(input_path), str(temp_copy))
                
                import pythoncom  # type: ignore[import-not-found]
                import win32com.client  # type: ignore[import-not-found]

                pythoncom.CoInitialize()
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(str(temp_copy.resolve()), ReadOnly=True)
                
                # Small delay to ensure file is fully loaded
                time.sleep(0.2)
                
                doc.ExportAsFixedFormat(str(output_pdf.resolve()), 17)
                logger.info("Output DOCX preview built: %s -> %s", input_path, output_pdf)
                
                # Success, break retry loop
                break
                
            except Exception as exc:  # noqa: BLE001
                error_str = str(exc)
                # Check if error is "file in use"
                if "используется другим приложением" in error_str or "used by another" in error_str.lower():
                    logger.warning("File is locked by another application (likely Word). Skipping PDF preview for %s.", input_path)
                    # Return original path so UI can show DOCX preview instead
                    return input_path
                
                logger.warning("Attempt %d to convert DOCX failed: %s", attempt + 1, exc)
                if attempt < 2:
                    time.sleep(1)  # Wait before retry
                else:
                    logger.exception("Failed to convert DOCX output to PDF: %s", input_path)
                    # If all retries fail, return original path to avoid crashing the whole preview generation
                    return input_path
            finally:
                # Cleanup Word instance for this attempt
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                    doc = None
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                    word = None
                try:
                    import pythoncom  # type: ignore[import-not-found]
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                # Cleanup temp copy
                try:
                    if temp_copy.exists():
                        temp_copy.unlink(missing_ok=True)
                    if temp_dir.exists() and not any(temp_dir.iterdir()):
                        temp_dir.rmdir()
                except Exception:
                    pass

        if not output_pdf.exists():
            logger.warning("Preview PDF was not created, falling back to original file: %s", input_path)
            return input_path
            
        self._temp_pdfs.add(output_pdf)
        return output_pdf

    def convert_xlsx_to_pdf(self, input_path: Path) -> Path:
        import uuid
        output_pdf = self.preview_dir / f"{uuid.uuid4().hex[:8]}_{input_path.stem}.pdf"
        excel = None
        wb = None

        import shutil
        import tempfile
        import time
        import uuid

        # Validate input file exists
        if not input_path.exists():
            logger.warning("XLSX preview input not found: %s, returning original path", input_path)
            return input_path

        # Retry loop to handle transient issues
        for attempt in range(3):
            temp_dir = Path(tempfile.mkdtemp(prefix="contract_xlsx_preview_"))
            unique_name = f"{uuid.uuid4().hex[:8]}_preview.xlsx"
            temp_copy = temp_dir / unique_name

            try:
                # Copy file to temp location with simple name
                shutil.copy2(str(input_path), str(temp_copy))

                import pythoncom  # type: ignore[import-not-found]
                import win32com.client  # type: ignore[import-not-found]

                pythoncom.CoInitialize()
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(str(temp_copy.resolve()), ReadOnly=True)

                # Small delay to ensure file is fully loaded
                time.sleep(0.3)

                wb.ExportAsFixedFormat(0, str(output_pdf.resolve()))
                logger.info("Output XLSX preview built: %s -> %s", input_path, output_pdf)

                # Success, break retry loop
                break

            except Exception as exc:  # noqa: BLE001
                error_str = str(exc)
                # Check if error is "file in use" or "used by another"
                if "используется другим приложением" in error_str.lower() or "used by another" in error_str.lower() or "открыт" in error_str.lower():
                    logger.warning("File is locked by another application (likely Excel). Skipping PDF preview for %s.", input_path)
                    return input_path

                logger.warning("Attempt %d to convert XLSX failed: %s", attempt + 1, exc)
                if attempt < 2:
                    time.sleep(1)  # Wait before retry
                else:
                    # All attempts failed - return original file instead of crashing
                    logger.warning("All XLSX to PDF conversion attempts failed, returning original file: %s", input_path)
                    return input_path
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                    wb = None
                if excel is not None:
                    try:
                        excel.Quit()
                    except Exception:
                        pass
                    excel = None
                try:
                    import pythoncom  # type: ignore[import-not-found]
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                # Cleanup temp copy
                try:
                    if temp_copy.exists():
                        temp_copy.unlink(missing_ok=True)
                    if temp_dir.exists() and not any(temp_dir.iterdir()):
                        temp_dir.rmdir()
                except Exception:
                    pass

        if not output_pdf.exists():
            logger.warning("Preview PDF was not created, falling back to original file: %s", input_path)
            return input_path

        self._temp_pdfs.add(output_pdf)
        return output_pdf

    def cleanup(self) -> None:
        for path in list(self._temp_pdfs):
            try:
                path.unlink(missing_ok=True)
            except Exception:  # noqa: BLE001
                logger.exception("Failed to cleanup output preview PDF: %s", path)
            finally:
                self._temp_pdfs.discard(path)

    def clear_all(self) -> None:
        """Удаляет все файлы из директории предпросмотра."""
        import gc
        import time

        # Force garbage collection to release COM-locked files
        gc.collect()
        time.sleep(0.2)

        if not self.preview_dir.exists():
            return

        locked_count = 0
        for path in self.preview_dir.iterdir():
            if path.is_file():
                try:
                    path.unlink(missing_ok=True)
                except PermissionError:
                    locked_count += 1
                    logger.debug("Preview file is locked, keeping: %s", path)
                except Exception:
                    logger.exception("Failed to delete preview file: %s", path)

        self._temp_pdfs.clear()

        if locked_count > 0:
            logger.info("Output preview directory cleared (except %d locked files): %s", locked_count, self.preview_dir)
        else:
            logger.info("Output preview directory cleared: %s", self.preview_dir)
