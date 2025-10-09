import json
from pathlib import Path

import pytest

from excel_compare.application.archive.use_cases import ArchiveComparisonUseCase
from excel_compare.domain.archive.entities import ArchiveFile, ArchiveComparisonRequest
from excel_compare.infrastructure.archive.file_repository import FileSystemArchiveRepository


@pytest.fixture
def repo(tmp_path: Path) -> FileSystemArchiveRepository:
    root = tmp_path / "history"
    return FileSystemArchiveRepository(root)


def test_archive_use_case_creates_run_directory(repo: FileSystemArchiveRepository, tmp_path: Path) -> None:
    use_case = ArchiveComparisonUseCase(repository=repo)
    request = ArchiveComparisonRequest(
        run_id="20241005_101500",
        inputs=[ArchiveFile(name="spectra.xls", content=b"spectra-bytes")],
        outputs=[ArchiveFile(name="comparison_all.xlsx", content=b"comp-bytes")],
    )

    receipt = use_case.execute(request)

    run_dir = tmp_path / "history" / "20241005_101500"
    assert run_dir.is_dir()
    assert (run_dir / "spectra.xls").read_bytes() == b"spectra-bytes"
    assert (run_dir / "comparison_all.xlsx").read_bytes() == b"comp-bytes"

    manifest_path = run_dir / "manifest.json"
    assert manifest_path.is_file()
    manifest = json.loads(manifest_path.read_text())
    assert manifest["run_id"] == "20241005_101500"
    assert {entry["name"] for entry in manifest["inputs"]} == {"spectra.xls"}
    assert {entry["name"] for entry in manifest["outputs"]} == {"comparison_all.xlsx"}

    assert receipt.run_id == "20241005_101500"
    assert receipt.location == run_dir


def test_archive_use_case_normalizes_run_id(repo: FileSystemArchiveRepository, tmp_path: Path) -> None:
    use_case = ArchiveComparisonUseCase(repository=repo)
    request = ArchiveComparisonRequest(
        run_id=" 2024/10/05 10:15:00 ",
        inputs=[ArchiveFile(name="lhs.xlsx", content=b"lhs")],
        outputs=[],
    )

    receipt = use_case.execute(request)

    expected_dir = tmp_path / "history" / "20241005_101500"
    assert expected_dir.is_dir()
    assert receipt.location == expected_dir

    manifest = json.loads((expected_dir / "manifest.json").read_text())
    assert manifest["run_id"] == "20241005_101500"
    assert manifest["inputs"] == [{"name": "lhs.xlsx", "bytes": len(b"lhs") }]
    assert manifest["outputs"] == []
