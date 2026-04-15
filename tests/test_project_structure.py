from pathlib import Path


def test_required_files_exist() -> None:
    root = Path(__file__).resolve().parents[1]
    required = [
        "main.py",
        "requirements.txt",
        "README.md",
        "config/app_config.json",
        "config/logging.json",
        "app/models/enums.py",
        "app/models/data_models.py",
        "app/services/interfaces/reader_interface.py",
        "app/services/interfaces/classifier_interface.py",
        "app/services/interfaces/parsers_interface.py",
        "app/services/interfaces/validator_interface.py",
        "app/services/interfaces/generators_interface.py",
        "app/ui/main_window.py",
        "app/core/app_controller.py",
    ]
    for relative in required:
        assert (root / relative).exists(), f"Missing file: {relative}"
