"""Domain services implementing comparison rules."""
from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Mapping, Sequence

from .models import Discrepancy, NavIdentity, NavRecord
from .results import ValidationReport, ValidationSummary


class NavValidator:
    """Executes zero-tolerance NAV comparisons between inbound and authoritative data."""

    def __init__(self, decimal_tolerance: Decimal | None = None) -> None:
        if decimal_tolerance is None:
            decimal_tolerance = Decimal("0")
        self._tolerance = decimal_tolerance

    def compare(self, inbound: Sequence[NavRecord], authoritative: Sequence[NavRecord]) -> ValidationReport:
        inbound_map = self._to_map(inbound)
        authoritative_map = self._to_map(authoritative)

        mismatches: list[Discrepancy] = []
        unauthorized: list[Discrepancy] = []
        missing: list[Discrepancy] = []
        duplicates: list[Discrepancy] = []

        duplicate_counts = self._detect_duplicates(inbound)
        duplicate_counts.update(self._detect_duplicates(authoritative))
        for identity, count in duplicate_counts.items():
            duplicates.append(
                Discrepancy(
                    identity=identity,
                    inbound=inbound_map.get(identity),
                    authoritative=authoritative_map.get(identity),
                    issue_type="duplicate",
                    message=f"{count} records share the same identity",
                )
            )

        for identity, auth_record in authoritative_map.items():
            inbound_record = inbound_map.get(identity)
            if inbound_record is None:
                missing.append(
                    Discrepancy(
                        identity=identity,
                        inbound=None,
                        authoritative=auth_record,
                        issue_type="missing_in_inbound",
                        message="No inbound NAV for authoritative record",
                    )
                )
                continue
            if not self._values_equal(inbound_record.nav, auth_record.nav):
                mismatches.append(
                    Discrepancy(
                        identity=identity,
                        inbound=inbound_record,
                        authoritative=auth_record,
                        issue_type="wrong_number",
                        message=f"Inbound NAV {inbound_record.nav} differs from authoritative {auth_record.nav}",
                    )
                )

        for identity, inbound_record in inbound_map.items():
            if identity not in authoritative_map:
                unauthorized.append(
                    Discrepancy(
                        identity=identity,
                        inbound=inbound_record,
                        authoritative=None,
                        issue_type="unauthorized_in_inbound",
                        message="Inbound NAV not present in authoritative data",
                    )
                )

        summary = ValidationSummary(
            total_inbound=len(inbound),
            total_authoritative=len(authoritative),
            wrong_numbers=len([d for d in mismatches if d.issue_type == "wrong_number"]),
            missing_in_inbound=len([d for d in missing if d.issue_type == "missing_in_inbound"]),
            unauthorized_in_inbound=len([d for d in unauthorized if d.issue_type == "unauthorized_in_inbound"]),
            duplicates=len(duplicates),
            stale=0,
            generated_at=datetime.utcnow(),
        )

        return ValidationReport(
            summary=summary,
            mismatches=tuple(mismatches + missing + unauthorized),
            duplicates=tuple(duplicates),
            stale=tuple(),
        )

    @staticmethod
    def _to_map(records: Sequence[NavRecord]) -> Mapping[NavIdentity, NavRecord]:
        return {record.identity: record for record in records}

    @staticmethod
    def _detect_duplicates(records: Sequence[NavRecord]) -> Mapping[NavIdentity, int]:
        counts: dict[NavIdentity, int] = defaultdict(int)
        for record in records:
            counts[record.identity] += 1
        return {identity: count for identity, count in counts.items() if count > 1}

    def _values_equal(self, left: Decimal, right: Decimal) -> bool:
        return abs(left - right) <= self._tolerance
