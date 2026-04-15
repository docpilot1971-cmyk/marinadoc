from __future__ import annotations

from pathlib import Path

from app.core.config import AppConfig
from app.models import PartyType


class TemplateLoader:
    def __init__(self, config: AppConfig) -> None:
        self._config = config
        self._templates_dir = Path(config.paths.templates_dir)

    @property
    def templates_dir(self) -> Path:
        return self._templates_dir

    def resolve_act_template(self, executor_type: PartyType | None) -> Path:
        if executor_type == PartyType.IP:
            # Priority: v2 template with {variable} markers
            v2_template = self._config.templates.act_word_template_ip_v2 or self._config.templates.act_word_template_v2
            if v2_template:
                v2_path = self._templates_dir / v2_template
                if v2_path.exists():
                    return v2_path
            # Fallback: v1 template with #marker# format
            template_name = self._config.templates.act_word_template_ip or self._config.templates.act_word_template
            return self._resolve_existing(template_name, fallback_patterns=("act*ip*filled_test*.docx", "act*ip*.docx", "act*.docx"))
        # ORG executor
        # Priority: v2 template with {variable} markers
        v2_template = self._config.templates.act_word_template_v2
        if v2_template:
            v2_path = self._templates_dir / v2_template
            if v2_path.exists():
                return v2_path
        # Fallback: v1 template with #marker# format
        template_name = self._config.templates.act_word_template
        return self._resolve_existing(template_name, fallback_patterns=("act*org*org*filled_test*.docx", "act*org*org*.docx", "act*.docx"))

    def resolve_ks2_template(self, executor_type: PartyType | None = None) -> Path:
        """Resolve KS-2 template based on executor type."""
        if executor_type == PartyType.IP:
            template_name = self._config.templates.ks2_template_ip
        else:
            template_name = self._config.templates.ks2_template_org_org

        # Ищем специфичный шаблон (включая incoming/)
        result = self._resolve_existing(
            template_name or "",
            fallback_patterns=(
                "ks2*.xlsx",
                "*кс2*.xlsx",
                "*кс-2*.xlsx",
            ),
        )
        
        return result

    def resolve_ks3_template(self, executor_type: PartyType | None = None) -> Path:
        """Resolve KS-3 template based on executor type."""
        if executor_type == PartyType.IP:
            template_name = self._config.templates.ks3_template_ip
        else:
            template_name = self._config.templates.ks3_template_org_org

        # Ищем специфичный шаблон (включая incoming/)
        result = self._resolve_existing(
            template_name or "",
            fallback_patterns=(
                "ks3*.xlsx",
                "*кс3*.xlsx",
                "*кс-3*.xlsx",
            ),
        )
        
        return result

    def _resolve_existing(self, relative_name: str, fallback_patterns: tuple[str, ...]) -> Path:
        direct = self._templates_dir / relative_name
        if direct.exists():
            return direct

        blank = self._templates_dir / "blank_templates"
        # Also check blank_templates/ for the exact template name
        blank_direct = blank / relative_name
        if blank_direct.exists():
            return blank_direct

        samples = self._templates_dir / "samples_filled"
        for base in (blank, samples, self._templates_dir):
            if not base.exists():
                continue
            for pattern in fallback_patterns:
                matches = sorted(
                    p for p in base.glob(pattern) if not p.name.startswith("~$")
                )
                if matches:
                    return matches[0]
        return direct
