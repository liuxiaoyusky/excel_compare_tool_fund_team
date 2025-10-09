"""Repository interfaces anchoring the domain layer."""
from __future__ import annotations

from typing import Protocol, Sequence

from .models import NavRecord


class InboundNavRepository(Protocol):
    """Provides NAV records ingested from inbound files."""

    def list_nav_records(self) -> Sequence[NavRecord]:
        ...


class AuthoritativeNavRepository(Protocol):
    """Provides authoritative NAV records."""

    def list_nav_records(self) -> Sequence[NavRecord]:
        ...
