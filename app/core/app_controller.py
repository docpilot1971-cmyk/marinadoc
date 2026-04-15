from __future__ import annotations

import logging
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QMessageBox

from app.core.config import AppConfig
from app.models import EstimateRow, ExtractionResult, RowType
from app.services.contract_document import ContractDocument
from app.services.document_preview_service import DocumentPreviewService
from app.services.external_editor_service import ExternalEditorService
from app.services.generated_document_manager import GeneratedDocumentManager
from app.services.interfaces import (
    IActWordGenerator,
    IContractReader,
    IContractTypeClassifier,
    IExtractionValidator,
    IHeaderParser,
    IKS2ExcelGenerator,
    IKS3ExcelGenerator,
    IObjectParser,
    IPartiesParser,
    IPeriodParser,
    ITableParser,
    ITotalsParser,
)
from app.services.output_preview_service import OutputPreviewService
from app.services.parsing_utils import try_parse_date
from app.ui.main_window import MainWindow

logger = logging.getLogger(__name__)


class AppController:
    def __init__(
        self,
        window: MainWindow,
        config: AppConfig,
        reader: IContractReader,
        classifier: IContractTypeClassifier,
        header_parser: IHeaderParser,
        parties_parser: IPartiesParser,
        object_parser: IObjectParser,
        period_parser: IPeriodParser,
        table_parser: ITableParser,
        totals_parser: ITotalsParser,
        validator: IExtractionValidator,
        act_generator: IActWordGenerator,
        ks2_generator: IKS2ExcelGenerator,
        ks3_generator: IKS3ExcelGenerator,
        document_preview_service: DocumentPreviewService,
        output_preview_service: OutputPreviewService,
        external_editor_service: ExternalEditorService,
        generated_document_manager: GeneratedDocumentManager,
    ) -> None:
        self.window = window
        self.config = config
        self.reader = reader
        self.classifier = classifier
        self.header_parser = header_parser
        self.parties_parser = parties_parser
        self.object_parser = object_parser
        self.period_parser = period_parser
        self.table_parser = table_parser
        self.totals_parser = totals_parser
        self.validator = validator
        self.act_generator = act_generator
        self.ks2_generator = ks2_generator
        self.ks3_generator = ks3_generator
        self.document_preview_service = document_preview_service
        self.output_preview_service = output_preview_service
        self.external_editor_service = external_editor_service
        self.generated_document_manager = generated_document_manager

        self.current_contract_path: Path | None = None
        self.current_contract: ContractDocument | None = None
        self.current_result = ExtractionResult()

        self._connect_signals()

    def _connect_signals(self) -> None:
        self.window.load_requested.connect(self.on_load_contract)
        self.window.recognize_requested.connect(self.on_recognize)
        self.window.generate_requested.connect(self.on_generate_preview)
        self.window.open_editor_requested.connect(self.on_open_for_edit)
        self.window.refresh_preview_requested.connect(self.on_refresh_preview)
        self.window.save_requested.connect(self.on_save)
        self.window.output_tab_changed.connect(self.on_output_tab_changed)

    def on_load_contract(self) -> None:
        # Защита от повторной загрузки
        if self.current_contract_path is not None:
            reply = QMessageBox.question(
                self.window,
                "Договор уже загружен",
                f"Договор уже загружен: {self.current_contract_path.name}\n\n"
                f"Загрузить новый? Текущие данные будут потеряны.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
            # Сбросить текущий договор
            self._clear_current_contract()

        # Очистка кеша и временных файлов перед загрузкой нового договора
        # Это предотвращает ошибки блокировки файлов и конфликты путей
        try:
            self.output_preview_service.clear_all()
            self.document_preview_service.cleanup_all()
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to cleanup previews before loading: %s", exc)

        file_path_str, _ = QFileDialog.getOpenFileName(
            self.window,
            "Выберите файл договора",
            "",
            "Contract files (*.docx *.doc);;All files (*.*)",
        )
        if not file_path_str:
            return

        self.current_contract_path = Path(file_path_str)
        try:
            # Загрузить DOCX для предпросмотра (по умолчанию режим DOCX)
            self.window.set_source_docx(self.current_contract_path)
            # Также создать PDF preview для переключения режимов
            try:
                pdf_path = self.document_preview_service.preview_document(self.current_contract_path)
                self.window.set_source_pdf(pdf_path)
            except Exception as pdf_exc:  # noqa: BLE001
                logger.warning("Failed to create PDF preview for source contract: %s", pdf_exc)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to build source contract preview.")
            self.window.set_source_preview_error(f"Ошибка preview исходного договора: {exc}")
            QMessageBox.warning(self.window, "Preview Error", str(exc))
            return

        try:
            self.current_contract = self.reader.read(self.current_contract_path)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to read contract.")
            QMessageBox.critical(self.window, "Read Error", str(exc))
            self.current_contract_path = None
            return

        self.window.set_status(f"Загружен: {self.current_contract_path.name}")

    def _clear_current_contract(self) -> None:
        """Очистить данные текущего договора и сбросить интерфейс."""
        self.current_contract_path = None
        self.current_contract = None
        self.current_result = ExtractionResult()
        
        # Очистить временные файлы генерации
        try:
            self.generated_document_manager.cleanup()
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to cleanup generated documents: %s", exc)
        
        self.window.hide_recognized_data()
        self.window.set_status("Готово. Загрузите договор.")

    def on_recognize(self) -> None:
        if not self.current_contract:
            QMessageBox.warning(self.window, "Нет договора", "Сначала загрузите договор.")
            return

        try:
            classification = self.classifier.classify(self.current_contract)
            result = ExtractionResult()
            result.document = self.header_parser.parse(self.current_contract)
            result.customer, result.executor = self.parties_parser.parse(self.current_contract)
            result.customer.type = classification.customer_type
            result.executor.type = classification.executor_type
            result.object_data = self.object_parser.parse(self.current_contract)
            result.period = self.period_parser.parse(self.current_contract)
            result.rows = self.table_parser.parse(self.current_contract, classification.table_grouping_mode)
            result.totals = self.totals_parser.parse(self.current_contract, result.rows)
            if result.document.act_date is None:
                result.document.act_date = date.today()
            result = self.validator.validate(result)
            self.current_result = result
            self.window.fill_form(result)
            self.window.set_status(f"Распознано. Статус: {result.validation_status} | Строк: {len(result.rows)}")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to recognize contract.")
            QMessageBox.critical(self.window, "Ошибка распознавания", str(exc))

    def on_generate_preview(self) -> None:
        self._apply_form_to_result()
        try:
            if hasattr(self.act_generator, "set_source_contract_path"):
                self.act_generator.set_source_contract_path(self.current_contract_path)
            out_paths = self.generated_document_manager.prepare_output_paths(self.current_result.document.contract_number)
            generated = {
                "act": self.act_generator.generate(self.current_result, out_paths["act"]),
                "ks2": self.ks2_generator.generate(self.current_result, out_paths["ks2"]),
                "ks3": self.ks3_generator.generate(self.current_result, out_paths["ks3"]),
            }
            for key, path in generated.items():
                self.generated_document_manager.set_generated_file(key, path)
                pdf = self.output_preview_service.build_preview(path)
                self.generated_document_manager.set_preview_file(key, pdf)
                self.window.set_output_pdf(key, pdf)
            self._show_current_output_preview()
            self.window.set_status("Preview сформирован. Можно открыть документ для ручного редактирования.")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to generate output previews.")
            QMessageBox.critical(self.window, "Generation Error", str(exc))

    def on_open_for_edit(self) -> None:
        key = self.window.selected_output_key()
        target = self.generated_document_manager.get_generated_file(key)
        if target is None:
            QMessageBox.information(self.window, "Нет документа", "Сначала сформируйте preview.")
            return
        try:
            self.external_editor_service.open_for_edit(target)
            self.window.set_status(f"Opened for editing: {target.name}")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to open external editor.")
            QMessageBox.critical(self.window, "Editor Error", str(exc))

    def on_refresh_preview(self) -> None:
        key = self.window.selected_output_key()
        target = self.generated_document_manager.get_generated_file(key)
        if target is None:
            QMessageBox.information(self.window, "Нет документа", "Сначала сформируйте preview.")
            return
        try:
            pdf = self.output_preview_service.build_preview(target)
            self.generated_document_manager.set_preview_file(key, pdf)
            self.window.set_output_pdf(key, pdf)
            self.window.set_status(f"Preview обновлен: {key.upper()}")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to refresh output preview.")
            QMessageBox.critical(self.window, "Preview Refresh Error", str(exc))

    def on_save(self) -> None:
        target_dir_str = QFileDialog.getExistingDirectory(self.window, "Select Destination Folder")
        if not target_dir_str:
            return
        target_dir = Path(target_dir_str)
        try:
            saved = self.generated_document_manager.save_final(target_dir)
            if not saved:
                QMessageBox.information(self.window, "Нет документов", "Сначала сформируйте preview.")
                return
            self.window.set_status(f"Saved {len(saved)} file(s) to {target_dir}")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to save final documents.")
            QMessageBox.critical(self.window, "Save Error", str(exc))

    def on_output_tab_changed(self, _key: str) -> None:
        self._show_current_output_preview()

    def _show_current_output_preview(self) -> None:
        key = self.window.selected_output_key()
        pdf = self.generated_document_manager.get_preview_file(key)
        if pdf is None:
            self.window.set_output_message(key, "Preview еще не сформирован.")
            return
        self.window.set_output_pdf(key, pdf)

    def _apply_form_to_result(self) -> None:
        form_data = self.window.read_form()

        self.current_result.document.contract_number = form_data.get("contract_number") or None
        self.current_result.document.contract_date = self._try_parse_date(form_data.get("contract_date", ""))
        self.current_result.document.document_city = form_data.get("document_city") or None
        self.current_result.document.act_date = self._try_parse_date(form_data.get("act_date", "")) or date.today()

        self.current_result.customer.full_name = form_data.get("customer_name_full") or None
        self.current_result.customer.representative_name = form_data.get("customer_representative_name") or None
        self.current_result.customer.representative_position = form_data.get("customer_representative_position") or None
        self.current_result.customer.representative_basis = form_data.get("customer_representative_basis") or None
        self.current_result.customer.inn = form_data.get("customer_inn") or None
        self.current_result.customer.kpp = form_data.get("customer_kpp") or None
        self.current_result.customer.ogrn = form_data.get("customer_ogrn") or None
        self.current_result.customer.ogrnip = form_data.get("customer_ogrnip") or None
        self.current_result.customer.address = form_data.get("customer_address") or None
        self.current_result.customer.bank_name = form_data.get("customer_bank_name") or None
        self.current_result.customer.rs = form_data.get("customer_rs") or None
        self.current_result.customer.ks = form_data.get("customer_ks") or None
        self.current_result.customer.bik = form_data.get("customer_bik") or None

        self.current_result.executor.full_name = form_data.get("executor_name_full") or None
        self.current_result.executor.representative_name = form_data.get("executor_representative_name") or None
        self.current_result.executor.representative_position = form_data.get("executor_representative_position") or None
        self.current_result.executor.representative_basis = form_data.get("executor_representative_basis") or None
        self.current_result.executor.inn = form_data.get("executor_inn") or None
        self.current_result.executor.kpp = form_data.get("executor_kpp") or None
        self.current_result.executor.ogrn = form_data.get("executor_ogrn") or None
        self.current_result.executor.ogrnip = form_data.get("executor_ogrnip") or None
        self.current_result.executor.address = form_data.get("executor_address") or None
        self.current_result.executor.bank_name = form_data.get("executor_bank_name") or None
        self.current_result.executor.rs = form_data.get("executor_rs") or None
        self.current_result.executor.ks = form_data.get("executor_ks") or None
        self.current_result.executor.bik = form_data.get("executor_bik") or None
        self.current_result.executor.passport = form_data.get("executor_passport") or None
        self.current_result.executor.registration = form_data.get("executor_registration") or None
        self.current_result.executor.tax_office = form_data.get("executor_tax_office") or None

        self.current_result.object_data.object_name = form_data.get("object_name") or None
        self.current_result.object_data.object_address = form_data.get("object_address") or None
        self.current_result.object_data.object_inventory_no = form_data.get("object_inventory_no") or None
        self.current_result.object_data.object_cadastral_no = form_data.get("object_cadastral_no") or None

        self.current_result.period.work_start_date = self._try_parse_date(form_data.get("work_start_date", ""))
        self.current_result.period.work_end_date_plan = self._try_parse_date(form_data.get("work_end_date_plan", ""))
        self.current_result.period.work_end_date_fact = self._try_parse_date(form_data.get("work_end_date_fact", ""))
        self.current_result.period.reporting_period = form_data.get("reporting_period") or None

        self.current_result.totals.works_total = self._try_parse_decimal(form_data.get("works_total", ""))
        self.current_result.totals.materials_total = self._try_parse_decimal(form_data.get("materials_total", ""))
        self.current_result.totals.transport_total = self._try_parse_decimal(form_data.get("transport_total", ""))
        self.current_result.totals.travel_total = self._try_parse_decimal(form_data.get("travel_total", ""))
        self.current_result.totals.total_without_vat = self._try_parse_decimal(form_data.get("total_without_vat", ""))
        self.current_result.totals.vat_rate = self._try_parse_decimal(form_data.get("vat_rate", ""))
        self.current_result.totals.vat_amount = self._try_parse_decimal(form_data.get("vat_amount", ""))
        self.current_result.totals.total_with_vat = self._try_parse_decimal(form_data.get("total_with_vat", ""))

        self.current_result.rows = self._map_rows_from_ui(form_data.get("rows", []))
        self.current_result = self.validator.validate(self.current_result)

    @staticmethod
    def _map_rows_from_ui(raw_rows: list[dict[str, str]]) -> list[EstimateRow]:
        mapped: list[EstimateRow] = []
        default_group_mode = EstimateRow().row_grouping_mode
        for idx, row in enumerate(raw_rows):
            row_type_raw = row.get("row_type", "ITEM")
            row_type = RowType.ITEM
            try:
                row_type = RowType(row_type_raw)
            except ValueError:
                pass
            mapped.append(
                EstimateRow(
                    row_grouping_mode=default_group_mode,
                    row_type=row_type,
                    row_section=row.get("row_section") or None,
                    row_number=row.get("row_number") or None,
                    row_name=row.get("row_name") or None,
                    row_unit=row.get("row_unit") or None,
                    row_quantity=AppController._try_parse_decimal_nullable(row.get("row_quantity", "")),
                    row_price=AppController._try_parse_decimal_nullable(row.get("row_price", "")),
                    row_amount=AppController._try_parse_decimal_nullable(row.get("row_amount", "")),
                    row_completion_date=AppController._try_parse_date(row.get("row_completion_date", "")),
                    row_sort_index=idx,
                )
            )
        return mapped

    @staticmethod
    def _try_parse_date(value: str) -> date | None:
        return try_parse_date(value)

    @staticmethod
    def _try_parse_decimal(value: str) -> Decimal:
        value = value.strip()
        if not value:
            return Decimal("0")
        try:
            return Decimal(value.replace(",", "."))
        except InvalidOperation:
            return Decimal("0")

    @staticmethod
    def _try_parse_decimal_nullable(value: str) -> Decimal | None:
        value = value.strip()
        if not value:
            return None
        try:
            return Decimal(value.replace(",", "."))
        except InvalidOperation:
            return None
