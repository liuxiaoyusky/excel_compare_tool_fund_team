"""HSBC (authoritative) Excel parser producing canonical NAV records."""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Sequence

import pandas as pd

from nav_checker.config import SETTINGS
from nav_checker.domain.models import NavIdentity, NavRecord
from nav_checker.infrastructure.parsing.utils import (
    compute_file_hash,
    ensure_bytes,
    extract_currency,
    extract_nav_date,
    extract_share_class,
    parse_decimal,
)
from nav_checker.infrastructure.storage.mapping_store import load_mapping

DEFAULT_SHEET_NAME = "HSBC Position Appraisal Report"


def _list_sheets(source: BytesIO | Path) -> list[str]:
    xls = pd.ExcelFile(source, engine="openpyxl")
    return xls.sheet_names


def _pick_sheet(source: BytesIO | Path, preferred: str) -> str:
    sheets = _list_sheets(source)
    if not sheets:
        raise ValueError("HSBC workbook has no sheets")
    if preferred in sheets:
        return preferred
    lower_map = {name.lower(): name for name in sheets}
    if preferred.lower() in lower_map:
        return lower_map[preferred.lower()]
    for name in sheets:
        if preferred.lower() in name.lower():
            return name
    return sheets[0]


def read_hsbc_raw(source: BytesIO | Path) -> pd.DataFrame:
    sheet_name = _pick_sheet(source, DEFAULT_SHEET_NAME)
    return pd.read_excel(
        source,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype=str,
        header=12,
    )


def normalize_hsbc(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    mapping = load_mapping()

    for col in ("Isin", "Ticker"):
        if col in work.columns and "Security ID" in work.columns:
            mask = work[col].isna() | (work[col].astype(str).str.strip() == "")
            work.loc[mask, col] = work.loc[mask, "Security ID"].map(mapping)

    for col in ("Isin", "Ticker", "Security ID"):
        if col in work.columns:
            work[col] = work[col].astype(str).str.strip().str.upper()

    if "Isin" in work.columns:
        work["Isin"] = work["Isin"].astype(str).str.split().str[0]

    return work


def hsbc_to_records(
    source: BytesIO | Path,
    nav_date: date | None = None,
    currency: str | None = None,
    share_class: str | None = None,
) -> Sequence[NavRecord]:
    raw_bytes = ensure_bytes(source)
    dataframe = read_hsbc_raw(BytesIO(raw_bytes))
    normalized = normalize_hsbc(dataframe)
    nav_date = nav_date or extract_nav_date(normalized)
    currency = currency or extract_currency(normalized)
    share_class = share_class or extract_share_class(normalized)

    digest = compute_file_hash(raw_bytes)
    as_of = datetime.utcnow()
    mapping = load_mapping()

    records: list[NavRecord] = []
    for idx, row in normalized.iterrows():
        security_id = str(row.get("Security ID", "")).strip().upper()
        if not security_id:
            id_value = str(row.get("Isin") or row.get("Ticker") or "").strip().upper()
            security_id = mapping.get(id_value, id_value)
        if not security_id:
            continue
        nav_value = parse_decimal(row.get("Book Market Value"))
        identity = NavIdentity(
            instrument_id=security_id,
            nav_date=nav_date,
            currency=currency,
            share_class=share_class,
        )
        records.append(
            NavRecord(
                identity=identity,
                nav=nav_value,
                source="hsbc",
                file_hash=digest,
                as_of=as_of,
                lineage=f"row={idx}",
            )
        )
    return records
