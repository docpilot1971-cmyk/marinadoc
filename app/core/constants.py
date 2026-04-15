from pathlib import Path
import os
import sys


def get_resource_path(relative_path: str) -> Path:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Production: Path to the folder where the .exe is located
        base_path = Path(os.path.dirname(sys.executable))
    else:
        # Development: Path to the project root
        base_path = Path(__file__).resolve().parents[2]
    return base_path / relative_path


PROJECT_ROOT = get_resource_path(".")
CONFIG_DIR = get_resource_path("config")
TEMPLATES_DIR = get_resource_path("templates")
RESOURCES_DIR = get_resource_path("resources")

APP_CONFIG_FILE = CONFIG_DIR / "app_config.json"
LOGGING_CONFIG_FILE = CONFIG_DIR / "logging.json"

DEFAULT_OUTPUT_DIR_NAME = "output"
