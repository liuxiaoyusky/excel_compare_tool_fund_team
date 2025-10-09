"""Diff report generators for NAV discrepancies."""
from __future__ import annotations

import csv
import io
from typing import Sequence

from nav_checker.domain.models import Discrepancy
from nav_checker.domain.results import ValidationReport


def discrepancies_to_rows(discrepancies: Sequence[Discrepancy]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in discrepancies:
        identity = item.identity
        inbound_nav = item.inbound.nav if item.inbound else ""
        authoritative_nav = item.authoritative.nav if item.authoritative else ""
        rows.append(
            {
                "instrument_id": identity.instrument_id,
                "date": identity.nav_date.isoformat(),
                "currency": identity.currency,
                "share_class": identity.share_class,
                "issue_type": item.issue_type,
                "message": item.message,
                "inbound_nav": str(inbound_nav),
                "authoritative_nav": str(authoritative_nav),
            }
        )
    return rows


def render_csv(discrepancies: Sequence[Discrepancy]) -> bytes:
    rows = discrepancies_to_rows(discrepancies)
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=list(rows[0].keys()) if rows else [])
    if rows:
        writer.writeheader()
        writer.writerows(rows)
    return buffer.getvalue().encode("utf-8")


def render_html(report: ValidationReport) -> str:
    rows = discrepancies_to_rows(tuple(report.iter_all_discrepancies()))
    if not rows:
        return "<p>No discrepancies detected.</p>"
    header = "".join(f"<th>{col}</th>" for col in rows[0].keys())
    body_parts = []
    for row in rows:
        body_parts.append("<tr>" + "".join(f"<td>{value}</td>" for value in row.values()) + "</tr>")
    body_html = "".join(body_parts)
    return f"<table><thead><tr>{header}</tr></thead><tbody>{body_html}</tbody></table>"
