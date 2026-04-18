"""
Config persistence layer for FatturazioneXML tool.

Stores configuration in ~/Library/Application Support/FatturazioneXML/config.json.
"""

import json
from pathlib import Path

CONFIG_DIR = Path.home() / "Library" / "Application Support" / "FatturazioneXML"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_XLSM_PATH = str(Path.home() / "fatturazione" / "Database fatturazione 2026.xlsm")
DEFAULT_XML_OUTPUT_DIR = str(
    Path.home() / "fatturazione" / "Fatture 2026" / "Fatture 2026 XML"
)
DEFAULT_FILENAME_PREFIX = "IT01652160894_G"
DEFAULT_YEAR = 2026

KNOWN_KEYS = ("xlsm_path", "xml_output_dir", "filename_prefix", "year")


def _defaults() -> dict:
    return {
        "xlsm_path": DEFAULT_XLSM_PATH,
        "xml_output_dir": DEFAULT_XML_OUTPUT_DIR,
        "filename_prefix": DEFAULT_FILENAME_PREFIX,
        "year": DEFAULT_YEAR,
    }


def load_config() -> dict:
    """
    Load config from CONFIG_FILE. If the file doesn't exist, return defaults.
    Returns a dict with keys: xlsm_path, xml_output_dir, filename_prefix, year.
    Missing keys are filled in from defaults (forward-compatible).
    """
    config = _defaults()
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as fh:
                loaded = json.load(fh)
            # Merge: only known keys, loaded values override defaults
            for key in KNOWN_KEYS:
                if key in loaded:
                    config[key] = loaded[key]
        except (json.JSONDecodeError, OSError):
            pass  # Corrupt or unreadable file → return defaults
    return config


def save_config(config: dict) -> None:
    """
    Save config dict to CONFIG_FILE (creates CONFIG_DIR if needed).
    Only saves known keys (xlsm_path, xml_output_dir, filename_prefix, year).
    """
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    to_save = {key: config[key] for key in KNOWN_KEYS if key in config}
    with CONFIG_FILE.open("w", encoding="utf-8") as fh:
        json.dump(to_save, fh, indent=2)


def is_configured() -> bool:
    """
    Return True if CONFIG_FILE exists AND xlsm_path in config points to an existing file.
    """
    if not CONFIG_FILE.exists():
        return False
    config = load_config()
    return Path(config.get("xlsm_path", "")).exists()


def get_output_filename(config: dict, progressivo: str | int) -> str:
    """
    Compute the output filename (not full path) for the XML file.
    E.g. "IT01652160894_G48.xml" (prefix + progressivo + ".xml")
    """
    prefix = config.get("filename_prefix", DEFAULT_FILENAME_PREFIX)
    return f"{prefix}{str(progressivo)}.xml"


def get_output_path(config: dict, progressivo: str | int) -> Path:
    """
    Compute the full output path for the XML file.
    Creates the output directory if it doesn't exist.
    Returns: Path(xml_output_dir) / get_output_filename(config, progressivo)
    """
    output_dir = Path(config.get("xml_output_dir", DEFAULT_XML_OUTPUT_DIR))
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / get_output_filename(config, progressivo)
