"""Filesystem repository for archiving comparison runs."""
from __future__ import annotations

import json
import re
from pathlib import Path

from nav_checker.domain.archive.entities import (
    ArchiveComparisonRequest,
    ArchiveFile,
    ArchiveReceipt,
    iter_all_files,
)


def _normalize_run_id(run_id: str) -> str:
    if not run_id:
        return "run"
    digits = re.findall(r"\d", run_id)
    if len(digits) >= 14:
        date_part = "".join(digits[:8])
        time_part = "".join(digits[8:14])
        rest = "".join(digits[14:])
        normalized = f"{date_part}_{time_part}"
        if rest:
            normalized += rest
        return normalized
    sanitized = re.sub(r"[^0-9A-Za-z_-]+", "", run_id.strip())
    return sanitized or "run"


class FileSystemArchiveRepository:
    def __init__(self, root: Path) -> None:
        self._root = Path(root)

    def save_run(self, request: ArchiveComparisonRequest) -> ArchiveReceipt:
        normalized_run_id = _normalize_run_id(request.run_id)
        run_dir = self._root / normalized_run_id
        run_dir.mkdir(parents=True, exist_ok=True)

        for file in iter_all_files(request):
            self._write_file(run_dir, file)

        manifest_path = run_dir / "manifest.json"
        manifest = {
            "run_id": normalized_run_id,
            "inputs": [self._manifest_entry(file) for file in request.inputs],
            "outputs": [self._manifest_entry(file) for file in request.outputs],
        }
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

        return ArchiveReceipt(run_id=normalized_run_id, location=run_dir)

    @staticmethod
    def _write_file(run_dir: Path, archive_file: ArchiveFile) -> None:
        target = run_dir / archive_file.name
        target.write_bytes(archive_file.content)

    @staticmethod
    def _manifest_entry(archive_file: ArchiveFile) -> dict[str, object]:
        return {"name": archive_file.name, "bytes": len(archive_file.content)}
