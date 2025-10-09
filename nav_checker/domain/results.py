"""Domain-level results for NAV validation."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable, Sequence

from .models import Discrepancy, NavRecord


@dataclass(frozen=True)
class ValidationSummary:
    total_inbound: int
    total_authoritative: int
    wrong_numbers: int
    missing_in_inbound: int
    unauthorized_in_inbound: int
    duplicates: int
    stale: int
    generated_at: datetime


@dataclass(frozen=True)
class ValidationReport:
    summary: ValidationSummary
    mismatches: Sequence[Discrepancy] = field(default_factory=tuple)
    duplicates: Sequence[Discrepancy] = field(default_factory=tuple)
    stale: Sequence[Discrepancy] = field(default_factory=tuple)

    def has_issues(self) -> bool:
        return any(
            [
                self.summary.wrong_numbers,
                self.summary.missing_in_inbound,
                self.summary.unauthorized_in_inbound,
                self.summary.duplicates,
                self.summary.stale,
            ]
        )

    def iter_all_discrepancies(self) -> Iterable[Discrepancy]:
        yield from self.mismatches
        yield from self.duplicates
        yield from self.stale
