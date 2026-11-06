"""Application messages only; source documents and number formats are untouched."""
from __future__ import annotations

import json
from pathlib import Path

ENGLISH = json.loads(Path(__file__).with_name("translations_en.json").read_text(encoding="utf-8"))
_language = "ru"


def set_language(language: str) -> None:
    global _language
    if language not in ("ru", "en"):
        raise ValueError("Unsupported language")
    _language = language


def get_language() -> str:
    return _language


def tr(source: str) -> str:
    return ENGLISH.get(source, source) if _language == "en" else source
