"""Application-level DTOs for NAV validation."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

from nav_checker.domain.models import NavRecord
from nav_checker.domain.results import ValidationReport


@dataclass(slots=True, frozen=True)
class ValidationRequest:
    inbound_source: str
    authoritative_source: str
    inbound_files: Sequence[Path]
    run_id: str
    received_at: datetime


@dataclass(slots=True, frozen=True)
class ValidationResponse:
    report: ValidationReport
    normalized_inbound: Sequence[NavRecord]
    normalized_authoritative: Sequence[NavRecord]
