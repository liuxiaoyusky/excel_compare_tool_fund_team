from pathlib import Path
import json

from nav_checker.infrastructure.storage.mapping_store import load_mapping, save_mapping


def test_save_and_load_mapping(tmp_path: Path):
    path = tmp_path / "mapping_override.json"
    merged = save_mapping({"abc": "123"}, path=path)
    assert merged["ABC"] == "123"
    assert json.loads(path.read_text()) == {"ABC": "123"}

    loaded = load_mapping(path=path)
    assert loaded["ABC"] == "123"
