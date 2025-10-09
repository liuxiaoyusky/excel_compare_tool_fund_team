"""Archive domain entities for storing comparison runs."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence


@dataclass(frozen=True)
class ArchiveFile:
    name: str
    content: bytes


@dataclass(frozen=True)
class ArchiveComparisonRequest:
    run_id: str
    inputs: Sequence[ArchiveFile]
    outputs: Sequence[ArchiveFile]


@dataclass(frozen=True)
class ArchiveReceipt:
    run_id: str
    location: Path


def iter_all_files(request: ArchiveComparisonRequest) -> Iterable[ArchiveFile]:
    yield from request.inputs
    yield from request.outputs
