from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from app.core import AppController, load_app_config, setup_logging
from app.services.document_preview_service import DocumentPreviewService
from app.services.excel_template_processor import ExcelTemplateProcessor
from app.services.external_editor_service import ExternalEditorService
from app.services.generated_document_manager import GeneratedDocumentManager
from app.services.output_preview_service import OutputPreviewService
from app.services.stubs import (
    ActWordGeneratorStub,
    ContractReaderStub,
    ContractTypeClassifierStub,
    ExtractionValidatorStub,
    HeaderParserStub,
    KS2ExcelGeneratorStub,
    KS3ExcelGeneratorStub,
    ObjectParserStub,
    PartiesParserStub,
    PeriodParserStub,
    TableParserStub,
    TotalsParserStub,
)
from app.services.template_loader import TemplateLoader
from app.services.word_template_processor import WordTemplateProcessor
from app.ui import MainWindow


def build_app() -> tuple[QApplication, MainWindow, AppController]:
    setup_logging()
    config = load_app_config()
    template_loader = TemplateLoader(config)
    word_processor = WordTemplateProcessor()
    excel_processor = ExcelTemplateProcessor()
    document_preview_service = DocumentPreviewService()
    output_preview_service = OutputPreviewService(Path(config.paths.preview_dir))
    external_editor_service = ExternalEditorService()
    generated_document_manager = GeneratedDocumentManager()

    qt_app = QApplication(sys.argv)

    # Set application icon
    icon_path = Path(__file__).parent / "resources" / "app_icon.ico"
    if icon_path.exists():
        from PySide6.QtGui import QIcon
        qt_app.setWindowIcon(QIcon(str(icon_path)))
    qt_app.setApplicationName("MarinaDoc")
    qt_app.setOrganizationName("MarinaDoc")

    qt_app.aboutToQuit.connect(document_preview_service.cleanup_all)
    qt_app.aboutToQuit.connect(output_preview_service.cleanup)
    qt_app.aboutToQuit.connect(generated_document_manager.cleanup)
    window = MainWindow(title=config.window_title)

    # Also set window icon
    if icon_path.exists():
        from PySide6.QtGui import QIcon
        window.setWindowIcon(QIcon(str(icon_path)))

    controller = AppController(
        window=window,
        config=config,
        reader=ContractReaderStub(),
        classifier=ContractTypeClassifierStub(),
        header_parser=HeaderParserStub(),
        parties_parser=PartiesParserStub(),
        object_parser=ObjectParserStub(),
        period_parser=PeriodParserStub(),
        table_parser=TableParserStub(),
        totals_parser=TotalsParserStub(),
        validator=ExtractionValidatorStub(),
        act_generator=ActWordGeneratorStub(template_loader, word_processor),
        ks2_generator=KS2ExcelGeneratorStub(template_loader, excel_processor),
        ks3_generator=KS3ExcelGeneratorStub(template_loader, excel_processor),
        document_preview_service=document_preview_service,
        output_preview_service=output_preview_service,
        external_editor_service=external_editor_service,
        generated_document_manager=generated_document_manager,
    )
    return qt_app, window, controller


def main() -> int:
    qt_app, window, _controller = build_app()
    window.show()
    return qt_app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
