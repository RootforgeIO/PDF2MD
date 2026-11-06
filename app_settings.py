"""Per-user preferences, separate from documents and optional Marker state."""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from marker_install import user_data_dir

DEFAULTS = dict(language="ru", engine="auto", page_markers=True,
                page_separators=True, save_images=True, strip_headers=True, last_dir="")


def settings_path() -> Path:
    override = os.environ.get("PDF2MD_SETTINGS_FILE")
    return Path(override) if override else user_data_dir() / "settings.json"


def validated(data) -> dict:
    result = DEFAULTS.copy()
    if not isinstance(data, dict):
        return result
    for key, default in DEFAULTS.items():
        value = data.get(key)
        if type(value) is type(default):
            result[key] = value
    if result["language"] not in ("ru", "en"):
        result["language"] = "ru"
    if result["engine"] not in ("auto", "marker", "structured", "pymupdf"):
        result["engine"] = "auto"
    if result["last_dir"] and not Path(result["last_dir"]).is_absolute():
        result["last_dir"] = ""
    return result


def load_settings() -> dict:
    try:
        return validated(json.loads(settings_path().read_text(encoding="utf-8")))
    except (OSError, ValueError, UnicodeError):
        return DEFAULTS.copy()


def save_settings(data: dict) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix="settings-", suffix=".tmp", dir=path.parent)
    pending = Path(name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(validated(data), stream, ensure_ascii=False, indent=2)
        os.replace(pending, path)
    finally:
        pending.unlink(missing_ok=True)
