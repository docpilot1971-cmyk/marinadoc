from __future__ import annotations

import logging
from pathlib import Path

from app.models import ExtractionResult
from app.models.enums import PartyType
from app.services.excel_template_processor import ExcelTemplateProcessor
from app.services.interfaces import IActWordGenerator, IKS2ExcelGenerator, IKS3ExcelGenerator
from app.services.template_loader import TemplateLoader
from app.services.word_template_processor import WordTemplateProcessor

logger = logging.getLogger(__name__)


class ActWordGeneratorStub(IActWordGenerator):
    def __init__(self, template_loader: TemplateLoader, processor: WordTemplateProcessor) -> None:
        self.template_loader = template_loader
        self.processor = processor
        self.source_contract_path: Path | None = None

    def set_source_contract_path(self, path: Path | None) -> None:
        self.source_contract_path = path

    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        executor_type = data.executor.type if isinstance(data.executor.type, PartyType) else None
        template_path = self.template_loader.resolve_act_template(executor_type)
        result = self.processor.render(template_path, data, output_path, source_contract_path=self.source_contract_path)
        logger.info("Act generated from template: %s", template_path)
        return result


class KS2ExcelGeneratorStub(IKS2ExcelGenerator):
    def __init__(self, template_loader: TemplateLoader, processor: ExcelTemplateProcessor) -> None:
        self.template_loader = template_loader
        self.processor = processor

    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        executor_type = data.executor.type if isinstance(data.executor.type, PartyType) else None
        template_path = self.template_loader.resolve_ks2_template(executor_type)
        result = self.processor.render_ks2(template_path, data, output_path)
        logger.info("KS-2 generated from template: %s", template_path)
        return result


class KS3ExcelGeneratorStub(IKS3ExcelGenerator):
    def __init__(self, template_loader: TemplateLoader, processor: ExcelTemplateProcessor) -> None:
        self.template_loader = template_loader
        self.processor = processor

    def generate(self, data: ExtractionResult, output_path: Path) -> Path:
        executor_type = data.executor.type if isinstance(data.executor.type, PartyType) else None
        template_path = self.template_loader.resolve_ks3_template(executor_type)
        result = self.processor.render_ks3(template_path, data, output_path)
        logger.info("KS-3 generated from template: %s", template_path)
        return result
