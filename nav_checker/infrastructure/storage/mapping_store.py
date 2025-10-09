"""Storage helpers for security ID mappings."""
from __future__ import annotations

from pathlib import Path
import json
from typing import Any

from seg_mapping_config import sec_mapping


DEFAULT_PATH = Path(__file__).resolve().parents[2] / "mapping_override.json"


def _normalize_mapping(raw: dict[str, Any] | None) -> dict[str, str]:
    normalized: dict[str, str] = {}
    if not isinstance(raw, dict):
        return normalized
    for key, value in raw.items():
        if key is None:
            continue
        key_str = str(key).strip().upper()
        if not key_str:
            continue
        value_str = "" if value is None else str(value).strip().upper()
        normalized[key_str] = value_str
    return normalized


def load_mapping(path: Path | None = None) -> dict[str, str]:
    override_path = path or DEFAULT_PATH
    default_mapping = _normalize_mapping(sec_mapping)
    if not override_path.exists():
        return default_mapping
    try:
        data = json.loads(override_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default_mapping
    override = _normalize_mapping(data)
    default_mapping.update(override)
    return default_mapping


def save_mapping(mapping: dict[str, str], path: Path | None = None) -> dict[str, str]:
    override_path = path or DEFAULT_PATH
    normalized = _normalize_mapping(mapping)
    override_path.write_text(
        json.dumps(normalized, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    merged = dict(_normalize_mapping(sec_mapping))
    merged.update(normalized)
    return merged
