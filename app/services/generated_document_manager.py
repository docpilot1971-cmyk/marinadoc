from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path


class GeneratedDocumentManager:
    def __init__(self) -> None:
        self._base_dir = Path(tempfile.mkdtemp(prefix="contract_generated_"))
        self._session_dir: Path | None = None
        self._generated_files: dict[str, Path] = {}
        self._preview_files: dict[str, Path] = {}
        self._contract_number: str = "without_number"

    def prepare_output_paths(self, contract_number: str | None) -> dict[str, Path]:
        # Пересоздаём base_dir для каждой новой генерации
        import gc
        gc.collect()

        # Clean up old base_dir before creating new one
        if self._base_dir and self._base_dir.exists():
            try:
                shutil.rmtree(self._base_dir, ignore_errors=True)
            except Exception:
                pass

        self._base_dir = Path(tempfile.mkdtemp(prefix="contract_generated_"))
        
        self._contract_number = self._safe_name(contract_number or "without_number")
        self._session_dir = self._base_dir / self._contract_number
        if self._session_dir.exists():
            shutil.rmtree(self._session_dir, ignore_errors=True)
        self._session_dir.mkdir(parents=True, exist_ok=True)

        paths = {
            "act": self._session_dir / "temp_act.docx",
            "ks2": self._session_dir / "temp_ks2.xlsx",
            "ks3": self._session_dir / "temp_ks3.xlsx",
        }
        self._generated_files = paths.copy()
        self._preview_files = {}
        return paths

    def set_generated_file(self, key: str, path: Path) -> None:
        self._generated_files[key] = path

    def get_generated_file(self, key: str) -> Path | None:
        return self._generated_files.get(key)

    def generated_files(self) -> dict[str, Path]:
        return dict(self._generated_files)

    def set_preview_file(self, key: str, path: Path) -> None:
        self._preview_files[key] = path

    def get_preview_file(self, key: str) -> Path | None:
        return self._preview_files.get(key)

    def preview_files(self) -> dict[str, Path]:
        return dict(self._preview_files)

    def save_final(self, target_dir: Path) -> dict[str, Path]:
        target_dir.mkdir(parents=True, exist_ok=True)
        result: dict[str, Path] = {}
        naming = {
            "act": f"Акт_{self._contract_number}.docx",
            "ks2": f"КС2_{self._contract_number}.xlsx",
            "ks3": f"КС3_{self._contract_number}.xlsx",
        }
        for key, source in self._generated_files.items():
            if not source.exists():
                continue
            out = target_dir / naming[key]
            shutil.copy2(source, out)
            result[key] = out
        return result

    def cleanup(self) -> None:
        import gc
        import time
        
        # Перед удалением файлов пытаемся закрыть любые COM-объекты
        gc.collect()
        time.sleep(0.3)
        
        if self._base_dir.exists():
            shutil.rmtree(self._base_dir, ignore_errors=True)
        
        # Если ignore_errors не сработал, пробуем ещё раз с задержкой
        if self._base_dir.exists():
            time.sleep(1)
            try:
                shutil.rmtree(self._base_dir, ignore_errors=False)
            except Exception:
                pass  # Игнорируем ошибки при cleanup
        
        self._generated_files = {}
        self._preview_files = {}
        self._session_dir = None

    @staticmethod
    def _safe_name(value: str) -> str:
        cleaned = re.sub(r"[^\w\-.]+", "_", value, flags=re.UNICODE)
        return cleaned.strip("_") or "without_number"
