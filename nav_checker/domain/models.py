"""Domain models for NAV validation pipeline.

These dataclasses capture the canonical schema for normalized NAV records.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal


@dataclass(frozen=True)
class NavIdentity:
    """Unique identifier for an instrument's NAV on a specific date."""

    instrument_id: str
    nav_date: date
    currency: str
    share_class: str

    def key(self) -> tuple[str, date, str, str]:
        return (self.instrument_id, self.nav_date, self.currency, self.share_class)


@dataclass(frozen=True)
class NavRecord:
    """Normalized NAV record as produced by ingestion or fetched authoritatively."""

    identity: NavIdentity
    nav: Decimal
    source: str
    file_hash: str
    as_of: datetime
    lineage: str | None = None


@dataclass(frozen=True)
class Discrepancy:
    """Represents an actionable issue discovered during validation."""

    identity: NavIdentity
    inbound: NavRecord | None
    authoritative: NavRecord | None
    issue_type: str
    message: str
