"""Excel-backed repositories for NAV data."""
from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Sequence

from nav_checker.domain.models import NavRecord
from nav_checker.domain.repositories import (
    AuthoritativeNavRepository,
    InboundNavRepository,
)
from nav_checker.infrastructure.parsing.hsbc import hsbc_to_records
from nav_checker.infrastructure.parsing.spectra import spectra_to_records
from nav_checker.infrastructure.parsing.utils import ensure_bytes


class SpectraInboundRepository(InboundNavRepository):
    def __init__(self, source: BytesIO | Path | bytes, nav_date: date | None = None) -> None:
        self._source = ensure_bytes(source)
        self._nav_date = nav_date

    def list_nav_records(self) -> Sequence[NavRecord]:
        return spectra_to_records(BytesIO(self._source), nav_date=self._nav_date)


class HsbcAuthoritativeRepository(AuthoritativeNavRepository):
    def __init__(self, source: BytesIO | Path | bytes, nav_date: date | None = None) -> None:
        self._source = ensure_bytes(source)
        self._nav_date = nav_date

    def list_nav_records(self) -> Sequence[NavRecord]:
        return hsbc_to_records(BytesIO(self._source), nav_date=self._nav_date)
