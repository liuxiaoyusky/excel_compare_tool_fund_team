"""Archive application use cases."""
from __future__ import annotations

from dataclasses import dataclass

from nav_checker.domain.archive.entities import ArchiveComparisonRequest, ArchiveReceipt
from nav_checker.infrastructure.archive.file_repository import FileSystemArchiveRepository


@dataclass(slots=True)
class ArchiveComparisonUseCase:
    repository: FileSystemArchiveRepository

    def execute(self, request: ArchiveComparisonRequest) -> ArchiveReceipt:
        return self.repository.save_run(request)
