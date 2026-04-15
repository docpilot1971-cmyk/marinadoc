from __future__ import annotations

from decimal import Decimal, InvalidOperation

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDockWidget,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from app.models import EstimateRow, ExtractionResult
from app.ui.document_preview_widget import DocumentPreviewWidget


class MainWindow(QMainWindow):
    # Основные сигналы
    load_requested = Signal()
    recognize_requested = Signal()
    generate_requested = Signal()
    save_requested = Signal()
    open_editor_requested = Signal()
    refresh_preview_requested = Signal()
    output_tab_changed = Signal(str)
    # Новый сигнал для показа распознанных данных
    recognized_data_requested = Signal()

    ROW_COLUMNS = [
        "row_type",
        "row_section",
        "row_number",
        "row_name",
        "row_unit",
        "row_quantity",
        "row_price",
        "row_amount",
        "row_completion_date",
    ]

    TAB_KEYS = ["act", "ks2", "ks3"]

    def __init__(self, title: str = "Contract Report Generator") -> None:
        super().__init__()
        self.setWindowTitle(title)
        self.resize(1600, 1000)

        # Виджеты предпросмотра
        self.source_preview = DocumentPreviewWidget(enable_docx_mode=True)
        self.output_previews: dict[str, DocumentPreviewWidget] = {}

        # Панель редактирования (dock)
        self.edit_dock: QDockWidget | None = None
        self.edit_panel_visible = True

        # Данные для формы
        self.form_inputs: dict[str, QLineEdit] = {}
        self.rows_table: QTableWidget | None = None

        # Правая панель распознанных данных
        self.recognized_panel: QWidget | None = None
        self.recognized_placeholder: QLabel | None = None
        self.recognized_form_widget: QWidget | None = None

        self._build_toolbar()
        self._build_central_layout()
        self._build_edit_dock()
        self.statusBar().showMessage("Готово. Загрузите договор.")

    # ───────────────────── TOOLBAR ─────────────────────

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Основные действия")
        toolbar.setMovable(False)
        toolbar.setIconSize(toolbar.iconSize())
        self.addToolBar(toolbar)

        action_load = toolbar.addAction("📂 Загрузить договор")
        action_recognize = toolbar.addAction("🔍 Распознать")

        action_load.triggered.connect(self.load_requested.emit)
        action_recognize.triggered.connect(self.recognize_requested.emit)

        toolbar.addSeparator()

        self.generate_btn = toolbar.addAction("📄 Сформировать preview")
        self.edit_btn = toolbar.addAction("✏️ Открыть для редактирования")
        self.refresh_btn = toolbar.addAction("🔄 Обновить preview")
        self.save_btn = toolbar.addAction("💾 Сохранить")

        self.generate_btn.triggered.connect(self.generate_requested.emit)
        self.edit_btn.triggered.connect(self.open_editor_requested.emit)
        self.refresh_btn.triggered.connect(self.refresh_preview_requested.emit)
        self.save_btn.triggered.connect(self.save_requested.emit)

    # ───────────────────── CENTRAL LAYOUT ─────────────────────

    def _build_central_layout(self) -> None:
        """Строим центральную компоновку: лево (source) | право (recognized + output)."""
        main_splitter = QSplitter(Qt.Horizontal)

        # ── ЛЕВАЯ ПАНЕЛЬ: исходный договор ──
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(4, 4, 4, 4)

        # Переключатель режимов DOCX/PDF
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Режим просмотра:")
        mode_label.setStyleSheet("color: #555; font-size: 11px;")
        mode_layout.addWidget(mode_label)

        self.docx_mode_btn = QPushButton("DOCX")
        self.pdf_mode_btn = QPushButton("PDF")
        self.docx_mode_btn.setCheckable(True)
        self.pdf_mode_btn.setCheckable(True)
        self.docx_mode_btn.setChecked(True)
        self.docx_mode_btn.setFixedWidth(60)
        self.pdf_mode_btn.setFixedWidth(60)
        self.docx_mode_btn.clicked.connect(lambda: self._switch_source_mode("docx"))
        self.pdf_mode_btn.clicked.connect(lambda: self._switch_source_mode("pdf"))
        mode_layout.addWidget(self.docx_mode_btn)
        mode_layout.addWidget(self.pdf_mode_btn)
        mode_layout.addStretch()
        left_layout.addLayout(mode_layout)

        # Разделитель
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #ccc;")
        left_layout.addWidget(line)

        left_layout.addWidget(self.source_preview)
        main_splitter.addWidget(left_panel)

        # ── ПРАВАЯ ПАНЕЛЬ: распознанные данные + выход ──
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(4, 4, 4, 4)
        right_layout.setSpacing(4)

        # Верхняя часть: распознанные данные
        self.recognized_panel = QWidget()
        recognized_layout = QVBoxLayout(self.recognized_panel)
        recognized_layout.setContentsMargins(0, 0, 0, 0)
        recognized_layout.setSpacing(0)

        # Placeholder (показывается до распознавания)
        self.recognized_placeholder = QLabel(
            "📋 Распознанные данные\n\n"
            "Загрузите договор и нажмите «Распознать»\n"
            "для отображения извлечённых данных."
        )
        self.recognized_placeholder.setAlignment(Qt.AlignCenter)
        self.recognized_placeholder.setStyleSheet(
            "QLabel { color: #888; font-size: 14px; padding: 40px; "
            "background: #f9f9f9; border: 1px dashed #ccc; border-radius: 8px; }"
        )
        self.recognized_placeholder.setWordWrap(True)
        recognized_layout.addWidget(self.recognized_placeholder)

        # Форма + таблица (скрыты до распознавания)
        self.recognized_form_widget = self._create_form_and_table_widget(is_primary=True)
        self.recognized_form_widget.setVisible(False)
        recognized_layout.addWidget(self.recognized_form_widget)

        right_layout.addWidget(self.recognized_panel, stretch=1)

        # Разделитель
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        line2.setStyleSheet("color: #ccc;")
        right_layout.addWidget(line2)

        # Кнопка сворачивания панели редактирования
        self.toggle_edit_btn = QPushButton("▶ Развернуть редактирование")
        self._edit_panel_expanded = False  # По умолчанию скрыт
        self.toggle_edit_btn.setStyleSheet(
            "QPushButton { background: #e8e8e8; padding: 6px; border-radius: 4px; "
            "font-size: 12px; text-align: left; }"
            "QPushButton:hover { background: #d8d8d8; }"
        )
        self.toggle_edit_btn.clicked.connect(self._toggle_edit_panel)
        right_layout.addWidget(self.toggle_edit_btn)

        # Нижняя часть: табы выходных документов
        self.output_tabs = QTabWidget()
        self.output_tabs.currentChanged.connect(self._emit_output_tab_changed)

        self.output_previews["act"] = DocumentPreviewWidget()
        self.output_previews["ks2"] = DocumentPreviewWidget()
        self.output_previews["ks3"] = DocumentPreviewWidget()

        self.output_tabs.addTab(self.output_previews["act"], "📄 Акт")
        self.output_tabs.addTab(self.output_previews["ks2"], "📊 КС-2")
        self.output_tabs.addTab(self.output_previews["ks3"], "📊 КС-3")

        right_layout.addWidget(self.output_tabs, stretch=1)

        main_splitter.addWidget(right_panel)
        main_splitter.setSizes([600, 1000])

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.addWidget(main_splitter)
        self.setCentralWidget(container)

    # ───────────────────── FORM & TABLE WIDGET ─────────────────────

    def _create_form_and_table_widget(self, is_primary: bool = False) -> QWidget:
        """Создаёт виджет с формой и таблицей строк.
        
        Args:
            is_primary: Если True, поля добавляются в self.form_inputs (для правой панели)
        """
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # Scroll area для формы (чтобы не занимать всё пространство)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(350)
        scroll.setStyleSheet("QScrollArea { border: none; }")

        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setSpacing(3)
        form_layout.setContentsMargins(4, 4, 4, 4)

        local_inputs: dict[str, QLineEdit] = {}
        
        for key, label, placeholder in self._field_specs():
            input_widget = QLineEdit()
            input_widget.setPlaceholderText(placeholder)
            input_widget.setObjectName(key)
            input_widget.setMaximumHeight(24)
            if is_primary:
                self.form_inputs[key] = input_widget
            local_inputs[key] = input_widget
            form_layout.addRow(label, input_widget)

        scroll.setWidget(form_widget)
        layout.addWidget(scroll)

        # Таблица строк
        rows_table = QTableWidget(0, len(self.ROW_COLUMNS))
        rows_table.setHorizontalHeaderLabels(self.ROW_COLUMNS)
        rows_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        rows_table.setMaximumHeight(200)
        layout.addWidget(rows_table)

        # Кнопки управления строками
        btn_layout = QHBoxLayout()
        add_row_btn = QPushButton("➕ Добавить строку")
        remove_row_btn = QPushButton("➖ Удалить строку")
        # Connect buttons only for primary widget
        if is_primary:
            add_row_btn.clicked.connect(self._append_empty_row)
            remove_row_btn.clicked.connect(self._remove_selected_rows)
        add_row_btn.setMaximumHeight(28)
        remove_row_btn.setMaximumHeight(28)
        btn_layout.addWidget(add_row_btn)
        btn_layout.addWidget(remove_row_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Store reference to rows_table for primary widget
        if is_primary:
            self.rows_table = rows_table
            self._dock_inputs = local_inputs  # Store dock inputs for sync
        else:
            self._dock_inputs = local_inputs

        return widget

    # ───────────────────── EDIT DOCK ─────────────────────

    def _build_edit_dock(self) -> None:
        """Создаёт dock-панель редактирования (сворачиваемую)."""
        self.edit_dock = QDockWidget("Редактирование данных (Dock)")
        self.edit_dock.setAllowedAreas(Qt.BottomDockWidgetArea | Qt.RightDockWidgetArea)

        # Та же форма + таблица (но не primary)
        dock_widget = self._create_form_and_table_widget(is_primary=False)
        self.edit_dock.setWidget(dock_widget)

        # По умолчанию скрыт (данные в правой панели)
        self.edit_dock.setVisible(False)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.edit_dock)

    # ───────────────────── MODE SWITCHING ─────────────────────

    def _switch_source_mode(self, mode: str) -> None:
        """Переключить режим просмотра исходного документа."""
        if mode == "docx":
            self.docx_mode_btn.setChecked(True)
            self.pdf_mode_btn.setChecked(False)
            self.source_preview.set_mode("docx")
        elif mode == "pdf":
            self.docx_mode_btn.setChecked(False)
            self.pdf_mode_btn.setChecked(True)
            self.source_preview.set_mode("pdf")

    # ───────────────────── TOGGLE EDIT PANEL ─────────────────────

    def _toggle_edit_panel(self) -> None:
        """Свернуть/развернуть dock-панель редактирования."""
        if self.edit_dock is None:
            return

        self._edit_panel_expanded = not self._edit_panel_expanded

        if self._edit_panel_expanded:
            # Развернуть
            self.edit_dock.show()
            self.edit_dock.raise_()
            self.edit_dock.activateWindow()
            self.toggle_edit_btn.setText("▼ Свернуть редактирование")
        else:
            # Свернуть
            self.edit_dock.hide()
            self.toggle_edit_btn.setText("▶ Развернуть редактирование")

    # ───────────────────── OUTPUT TAB CHANGED ─────────────────────

    def _emit_output_tab_changed(self, index: int) -> None:
        if 0 <= index < len(self.TAB_KEYS):
            self.output_tab_changed.emit(self.TAB_KEYS[index])

    # ───────────────────── FIELD SPECS ─────────────────────

    @staticmethod
    def _field_specs() -> list[tuple[str, str, str]]:
        return [
            ("contract_number", "Номер договора", ""),
            ("contract_date", "Дата договора", "YYYY-MM-DD"),
            ("document_city", "Город", ""),
            ("act_date", "Дата акта", "YYYY-MM-DD"),
            ("customer_name_full", "Заказчик (полное)", ""),
            ("customer_representative_name", "Представитель заказчика", ""),
            ("customer_representative_position", "Должность заказчика", ""),
            ("customer_representative_basis", "Основание заказчика", ""),
            ("customer_inn", "ИНН заказчика", ""),
            ("customer_kpp", "КПП заказчика", ""),
            ("customer_ogrn", "ОГРН заказчика", ""),
            ("customer_ogrnip", "ОГРНИП заказчика", ""),
            ("customer_address", "Адрес заказчика", ""),
            ("customer_bank_name", "Банк заказчика", ""),
            ("customer_rs", "Р/с заказчика", ""),
            ("customer_ks", "К/с заказчика", ""),
            ("customer_bik", "БИК заказчика", ""),
            ("executor_name_full", "Подрядчик (полное)", ""),
            ("executor_representative_name", "Представитель подрядчика", ""),
            ("executor_representative_position", "Должность подрядчика", ""),
            ("executor_representative_basis", "Основание подрядчика", ""),
            ("executor_inn", "ИНН подрядчика", ""),
            ("executor_kpp", "КПП подрядчика", ""),
            ("executor_ogrn", "ОГРН подрядчика", ""),
            ("executor_ogrnip", "ОГРНИП подрядчика", ""),
            ("executor_passport", "Паспорт подрядчика", ""),
            ("executor_registration", "Регистрация ИП", ""),
            ("executor_tax_office", "ИФНС подрядчика", ""),
            ("executor_address", "Адрес подрядчика", ""),
            ("executor_bank_name", "Банк подрядчика", ""),
            ("executor_rs", "Р/с подрядчика", ""),
            ("executor_ks", "К/с подрядчика", ""),
            ("executor_bik", "БИК подрядчика", ""),
            ("object_name", "Объект", ""),
            ("object_address", "Адрес объекта", ""),
            ("object_inventory_no", "Инвентарный №", ""),
            ("object_cadastral_no", "Кадастровый №", ""),
            ("work_start_date", "Дата начала работ", "YYYY-MM-DD"),
            ("work_end_date_plan", "План окончания", "YYYY-MM-DD"),
            ("work_end_date_fact", "Факт окончания", "YYYY-MM-DD"),
            ("reporting_period", "Отчетный период", ""),
            ("works_total", "Итого работы", "0.00"),
            ("materials_total", "Итого материалы", "0.00"),
            ("transport_total", "Итого транспорт", "0.00"),
            ("travel_total", "Итого командировочные", "0.00"),
            ("total_without_vat", "Итого без НДС", "0.00"),
            ("vat_rate", "Ставка НДС", "20"),
            ("vat_amount", "Сумма НДС", "0.00"),
            ("total_with_vat", "Итого с НДС", "0.00"),
        ]

    # ───────────────────── ROW MANAGEMENT ─────────────────────

    def _append_empty_row(self) -> None:
        if self.rows_table is None:
            return
        row = self.rows_table.rowCount()
        self.rows_table.insertRow(row)
        for col in range(len(self.ROW_COLUMNS)):
            self.rows_table.setItem(row, col, QTableWidgetItem(""))

    def _remove_selected_rows(self) -> None:
        if self.rows_table is None:
            return
        selected_rows = sorted({idx.row() for idx in self.rows_table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            self.rows_table.removeRow(row)

    # ───────────────────── PUBLIC API ─────────────────────

    def set_source_docx(self, docx_path) -> None:  # noqa: ANN001
        """Загрузить исходный договор в режиме DOCX."""
        from pathlib import Path
        path = Path(docx_path) if not isinstance(docx_path, Path) else docx_path
        self.source_preview.load_docx(path)
        # Переключить на DOCX режим
        self.docx_mode_btn.setChecked(True)
        self.pdf_mode_btn.setChecked(False)
        self.source_preview.set_mode("docx")

    def set_source_pdf(self, pdf_path) -> None:  # noqa: ANN001
        """Загрузить PDF предпросмотр (в фоновом режиме, без переключения режима)."""
        self.source_preview.load_pdf(pdf_path)

    def set_source_preview_error(self, message: str) -> None:
        """Показать ошибку предпросмотра."""
        self.source_preview.show_message(message)

    def show_recognized_data(self) -> None:
        """Показать распознанные данные в правой панели."""
        if self.recognized_placeholder:
            self.recognized_placeholder.setVisible(False)
        if self.recognized_form_widget:
            self.recognized_form_widget.setVisible(True)
        # Also sync to dock if it exists
        self._sync_form_to_dock()

    def hide_recognized_data(self) -> None:
        """Скрыть распознанные данные (показать placeholder)."""
        if self.recognized_placeholder:
            self.recognized_placeholder.setVisible(True)
        if self.recognized_form_widget:
            self.recognized_form_widget.setVisible(False)
        # Also sync to dock
        self._sync_form_to_dock()

    def _sync_form_to_dock(self) -> None:
        """Sync form data from right panel to dock widget."""
        if not hasattr(self, '_dock_inputs') or not self._dock_inputs:
            return
        
        # Copy values from right panel form inputs to dock form inputs
        for key, dock_input in self._dock_inputs.items():
            right_input = self.form_inputs.get(key)
            if right_input is not None:
                dock_input.setText(right_input.text())

    def set_output_pdf(self, key: str, pdf_path) -> None:  # noqa: ANN001
        """Загрузить PDF выходного документа."""
        if key in self.output_previews:
            self.output_previews[key].load_pdf(pdf_path)

    def set_output_message(self, key: str, message: str) -> None:
        """Показать сообщение для выходного документа."""
        if key in self.output_previews:
            self.output_previews[key].show_message(message)

    def selected_output_key(self) -> str:
        """Получить ключ текущей вкладки."""
        idx = self.output_tabs.currentIndex()
        if 0 <= idx < len(self.TAB_KEYS):
            return self.TAB_KEYS[idx]
        return "act"

    # ───────────────────── FORM FILL/READ ─────────────────────

    def fill_form(self, result: ExtractionResult) -> None:
        """Заполнить форму распознанными данными."""
        values = {
            "contract_number": result.document.contract_number or "",
            "contract_date": str(result.document.contract_date or ""),
            "document_city": result.document.document_city or "",
            "act_date": str(result.document.act_date or ""),
            "customer_name_full": result.customer.full_name or "",
            "customer_representative_name": result.customer.representative_name or "",
            "customer_representative_position": result.customer.representative_position or "",
            "customer_representative_basis": result.customer.representative_basis or "",
            "customer_inn": result.customer.inn or "",
            "customer_kpp": result.customer.kpp or "",
            "customer_ogrn": result.customer.ogrn or "",
            "customer_ogrnip": result.customer.ogrnip or "",
            "customer_address": result.customer.address or "",
            "customer_bank_name": result.customer.bank_name or "",
            "customer_rs": result.customer.rs or "",
            "customer_ks": result.customer.ks or "",
            "customer_bik": result.customer.bik or "",
            "executor_name_full": result.executor.full_name or "",
            "executor_representative_name": result.executor.representative_name or "",
            "executor_representative_position": result.executor.representative_position or "",
            "executor_representative_basis": result.executor.representative_basis or "",
            "executor_inn": result.executor.inn or "",
            "executor_kpp": result.executor.kpp or "",
            "executor_ogrn": result.executor.ogrn or "",
            "executor_ogrnip": result.executor.ogrnip or "",
            "executor_passport": result.executor.passport or "",
            "executor_registration": result.executor.registration or "",
            "executor_tax_office": result.executor.tax_office or "",
            "executor_address": result.executor.address or "",
            "executor_bank_name": result.executor.bank_name or "",
            "executor_rs": result.executor.rs or "",
            "executor_ks": result.executor.ks or "",
            "executor_bik": result.executor.bik or "",
            "object_name": result.object_data.object_name or "",
            "object_address": result.object_data.object_address or "",
            "object_inventory_no": result.object_data.object_inventory_no or "",
            "object_cadastral_no": result.object_data.object_cadastral_no or "",
            "work_start_date": str(result.period.work_start_date or ""),
            "work_end_date_plan": str(result.period.work_end_date_plan or ""),
            "work_end_date_fact": str(result.period.work_end_date_fact or ""),
            "reporting_period": result.period.reporting_period or "",
            "works_total": str(result.totals.works_total),
            "materials_total": str(result.totals.materials_total),
            "transport_total": str(result.totals.transport_total),
            "travel_total": str(result.totals.travel_total),
            "total_without_vat": str(result.totals.total_without_vat),
            "vat_rate": str(result.totals.vat_rate),
            "vat_amount": str(result.totals.vat_amount),
            "total_with_vat": str(result.totals.total_with_vat),
        }
        for key, widget in self.form_inputs.items():
            widget.setText(values.get(key, ""))
        self._set_rows(result.rows)

        # Показать распознанные данные в правой панели
        self.show_recognized_data()

    def _set_rows(self, rows: list[EstimateRow]) -> None:
        if self.rows_table is None:
            return
        self.rows_table.setRowCount(0)
        for row_data in rows:
            row_idx = self.rows_table.rowCount()
            self.rows_table.insertRow(row_idx)
            values = [
                row_data.row_type.value,
                row_data.row_section or "",
                row_data.row_number or "",
                row_data.row_name or "",
                row_data.row_unit or "",
                self._decimal_to_str(row_data.row_quantity),
                self._decimal_to_str(row_data.row_price),
                self._decimal_to_str(row_data.row_amount),
                str(row_data.row_completion_date or ""),
            ]
            for col, value in enumerate(values):
                self.rows_table.setItem(row_idx, col, QTableWidgetItem(value))

    def read_form(self) -> dict[str, str | list[dict[str, str]]]:
        """Прочитать данные из формы."""
        payload: dict[str, str | list[dict[str, str]]] = {}
        for key, widget in self.form_inputs.items():
            payload[key] = widget.text().strip()
        payload["rows"] = self._read_rows()
        return payload

    def _read_rows(self) -> list[dict[str, str]]:
        if self.rows_table is None:
            return []
        rows: list[dict[str, str]] = []
        for row_idx in range(self.rows_table.rowCount()):
            row_data: dict[str, str] = {}
            for col_idx, col_name in enumerate(self.ROW_COLUMNS):
                item = self.rows_table.item(row_idx, col_idx)
                row_data[col_name] = item.text().strip() if item else ""
            if row_data["row_name"] or row_data["row_amount"] or row_data["row_number"]:
                rows.append(row_data)
        return rows

    @staticmethod
    def _decimal_to_str(value: Decimal | None) -> str:
        if value is None:
            return ""
        try:
            return str(value.quantize(Decimal("0.01")))
        except InvalidOperation:
            return str(value)

    def set_status(self, text: str) -> None:
        """Установить текст в статус-баре."""
        self.statusBar().showMessage(text)
