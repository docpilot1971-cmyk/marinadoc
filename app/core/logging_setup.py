from __future__ import annotations

import json
import logging
import logging.config
from pathlib import Path

from app.core.constants import LOGGING_CONFIG_FILE


def setup_logging(config_path: Path | None = None) -> None:
    target = config_path or LOGGING_CONFIG_FILE

    try:
        if target.exists():
            with target.open("r", encoding="utf-8") as f:
                config = json.load(f)
            logging.config.dictConfig(config)
            return
    except (PermissionError, OSError) as exc:
        # In PyInstaller .exe, temp files might have permission issues
        # Fall back to basic logging
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        )
        logging.getLogger(__name__).warning(
            "Could not load logging config from %s: %s. Using basic config.", target, exc
        )
        return

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )
