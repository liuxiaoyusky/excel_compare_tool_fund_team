"""Legacy compatibility wrapper pointing to the domain-driven NAV checker."""
from __future__ import annotations

from io import BytesIO
from typing import Any

from nav_checker import (
    NavValidationContext,
    NavValidator,
    SpectraInboundRepository,
    ValidateNavUseCase,
    HsbcAuthoritativeRepository,
)
from nav_checker.domain.results import ValidationReport
from nav_checker.presentation.diff_report import render_csv


def run_compare_from_sources(spectra_source: Any, hsbc_source: Any) -> dict[str, Any]:
    """Return a dictionary compatible with the legacy interface."""
    context = NavValidationContext(
        inbound_repository=SpectraInboundRepository(BytesIO(_ensure_bytes(spectra_source))),
        authoritative_repository=HsbcAuthoritativeRepository(BytesIO(_ensure_bytes(hsbc_source))),
        validator=NavValidator(),
    )
    use_case = ValidateNavUseCase(context)
    report, inbound_records, authoritative_records = use_case.execute()
    diff_csv = render_csv(tuple(report.iter_all_discrepancies()))
    return {
        "report": report,
        "inbound": inbound_records,
        "authoritative": authoritative_records,
        "diff_csv": diff_csv,
    }


def _ensure_bytes(source: Any) -> bytes:
    if isinstance(source, bytes):
        return source
    if hasattr(source, "read"):
        data = source.read()
        if isinstance(data, bytes):
            return data
    raise TypeError("Expected bytes or file-like object returning bytes")


def main(argv: list[str] | None = None) -> int:
    from nav_checker.cli import main as cli_main

    return cli_main(argv or [])


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
