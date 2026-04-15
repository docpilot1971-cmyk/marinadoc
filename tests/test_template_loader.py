import shutil
from pathlib import Path

from app.core.config import AppConfig, PathsConfig, TemplatesConfig
from app.models import PartyType
from app.services.template_loader import TemplateLoader


def _prepare_local_tmp_dir(name: str) -> Path:
    base = Path(__file__).resolve().parents[1] / ".pytest_local_tmp" / name
    if base.exists():
        shutil.rmtree(base)
    base.mkdir(parents=True, exist_ok=True)
    return base


def test_template_loader_resolves_direct_paths() -> None:
    tmp_path = _prepare_local_tmp_dir("loader_direct")
    templates_dir = tmp_path / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    (templates_dir / "act_org_org.docx").write_bytes(b"")
    (templates_dir / "act_org_ip.docx").write_bytes(b"")
    (templates_dir / "ks2.xlsx").write_bytes(b"")
    (templates_dir / "ks3.xlsx").write_bytes(b"")

    config = AppConfig(
        paths=PathsConfig(templates_dir=templates_dir, output_dir=tmp_path / "output", preview_dir=tmp_path / "preview"),
        templates=TemplatesConfig(
            act_word_template="act_org_org.docx",
            act_word_template_ip="act_org_ip.docx",
            ks2_template="ks2.xlsx",
            ks3_template="ks3.xlsx",
        ),
    )
    loader = TemplateLoader(config)

    assert loader.resolve_act_template(PartyType.ORG).name == "act_org_org.docx"
    assert loader.resolve_act_template(PartyType.IP).name == "act_org_ip.docx"
    assert loader.resolve_ks2_template().name == "ks2.xlsx"
    assert loader.resolve_ks3_template().name == "ks3.xlsx"


def test_template_loader_uses_fallback_from_incoming() -> None:
    tmp_path = _prepare_local_tmp_dir("loader_fallback")
    templates_dir = tmp_path / "templates"
    incoming = templates_dir / "incoming"
    incoming.mkdir(parents=True, exist_ok=True)
    (incoming / "act_org_ip_filled_test.docx").write_bytes(b"")
    (incoming / "ks2_test.xlsx").write_bytes(b"")
    (incoming / "rs3_test.xlsx").write_bytes(b"")

    config = AppConfig(
        paths=PathsConfig(templates_dir=templates_dir, output_dir=tmp_path / "output", preview_dir=tmp_path / "preview"),
        templates=TemplatesConfig(
            act_word_template="missing_org.docx",
            act_word_template_ip="missing_ip.docx",
            ks2_template="missing_ks2.xlsx",
            ks3_template="missing_ks3.xlsx",
        ),
    )
    loader = TemplateLoader(config)

    assert "ip" in loader.resolve_act_template(PartyType.IP).name.lower()
    assert loader.resolve_ks2_template().name.lower().startswith("ks2")
    assert loader.resolve_ks3_template().name.lower().startswith("rs3")
