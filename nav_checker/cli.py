"""Command-line entrypoint for NAV validation."""
from __future__ import annotations

import argparse
import sys
from datetime import date

from nav_checker.application.use_cases import NavValidationContext, ValidateNavUseCase
from nav_checker.domain.services import NavValidator
from nav_checker.infrastructure.parsing.utils import ensure_bytes
from nav_checker.infrastructure.repositories.excel_repositories import (
    HsbcAuthoritativeRepository,
    SpectraInboundRepository,
)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate inbound NAV files against HSBC authoritative data")
    parser.add_argument("spectra", type=str, help="Path to spectra Excel file")
    parser.add_argument("hsbc", type=str, help="Path to HSBC Excel file")
    parser.add_argument("--nav-date", type=str, help="Override NAV date (YYYY-MM-DD)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    nav_date = date.fromisoformat(args.nav_date) if args.nav_date else None

    inbound_repo = SpectraInboundRepository(args.spectra, nav_date=nav_date)
    authoritative_repo = HsbcAuthoritativeRepository(args.hsbc, nav_date=nav_date)

    context = NavValidationContext(
        inbound_repository=inbound_repo,
        authoritative_repository=authoritative_repo,
        validator=NavValidator(),
    )
    use_case = ValidateNavUseCase(context)
    report, inbound_records, authoritative_records = use_case.execute()

    print("Validation Summary")
    print("==================")
    summary = report.summary
    print(f"Inbound records: {summary.total_inbound}")
    print(f"Authoritative records: {summary.total_authoritative}")
    print(f"Wrong numbers: {summary.wrong_numbers}")
    print(f"Missing in inbound: {summary.missing_in_inbound}")
    print(f"Unauthorized inbound: {summary.unauthorized_in_inbound}")
    print(f"Duplicates: {summary.duplicates}")

    if report.has_issues():
        print("\nDiscrepancies detected:")
        for discrepancy in report.iter_all_discrepancies():
            identity = discrepancy.identity
            print(
                f"- {discrepancy.issue_type} for {identity.instrument_id} on {identity.nav_date}: {discrepancy.message}"
            )
    else:
        print("\nNo discrepancies detected.")

    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
