from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field

from app.core.constants import APP_CONFIG_FILE, DEFAULT_OUTPUT_DIR_NAME, get_resource_path


class PathsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")
    templates_dir: Path = get_resource_path("templates")
    output_dir: Path = Path(DEFAULT_OUTPUT_DIR_NAME)
    preview_dir: Path = Path("output/preview")


class TemplatesConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")
    # V1 templates: #marker# format
    act_word_template: str = "act_template.docx"
    act_word_template_ip: str | None = None
    # V2 templates: {variable} format
    act_word_template_v2: str = "act_org_org_filled_test_v2.docx"
    act_word_template_ip_v2: str = "act_org_ip_filled_test_v2.docx"
    
    # Excel KS-2 templates
    ks2_template_org_org: str = "ks2_org_org_test.xlsx"
    ks2_template_ip: str = "ks2_ip_test.xlsx"
    
    # Excel KS-3 templates
    ks3_template_org_org: str = "ks3_org_org_test.xlsx"
    ks3_template_ip: str = "ks3_ip_test.xlsx"


class AppConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")
    app_name: str = "Contract Report Generator"
    window_title: str = "Contract Report Generator"
    paths: PathsConfig = Field(default_factory=PathsConfig)
    templates: TemplatesConfig = Field(default_factory=TemplatesConfig)
    logging_config_path: Path = Path("config/logging.json")


def load_app_config(config_path: Path | None = None) -> AppConfig:
    target = config_path or APP_CONFIG_FILE

    try:
        if target.exists():
            with target.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return AppConfig.model_validate(data)
    except (PermissionError, OSError) as exc:
        # In PyInstaller .exe, temp files might have permission issues
        # Fall back to default config
        import logging
        logging.getLogger(__name__).warning(
            "Could not load app config from %s: %s. Using defaults.", target, exc
        )

    return AppConfig()
