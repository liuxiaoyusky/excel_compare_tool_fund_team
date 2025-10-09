"""Application services orchestrating the NAV validation workflow."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence

from nav_checker.domain.models import NavRecord
from nav_checker.domain.repositories import (
    AuthoritativeNavRepository,
    InboundNavRepository,
)
from nav_checker.domain.results import ValidationReport
from nav_checker.domain.services import NavValidator


@dataclass(slots=True)
class NavValidationContext:
    inbound_repository: InboundNavRepository
    authoritative_repository: AuthoritativeNavRepository
    validator: NavValidator


class ValidateNavUseCase:
    def __init__(self, context: NavValidationContext) -> None:
        self._context = context

    def execute(self) -> tuple[ValidationReport, Sequence[NavRecord], Sequence[NavRecord]]:
        inbound_records = self._context.inbound_repository.list_nav_records()
        authoritative_records = self._context.authoritative_repository.list_nav_records()
        report = self._context.validator.compare(inbound_records, authoritative_records)
        return report, inbound_records, authoritative_records
