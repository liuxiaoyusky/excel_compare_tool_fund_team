"""Spectra (inbound) Excel parser producing canonical NAV records."""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

from nav_checker.config import SETTINGS
from nav_checker.domain.models import NavIdentity, NavRecord
from nav_checker.infrastructure.storage.mapping_store import load_mapping
from nav_checker.infrastructure.parsing.utils import (
    compute_file_hash,
    ensure_bytes,
    extract_currency,
    extract_nav_date,
    extract_share_class,
    parse_decimal,
)

DEFAULT_SHEET_NAME = "Security Distribution"


def _list_sheets(source: BytesIO | Path) -> list[str]:
    if isinstance(source, Path):
        xls = pd.ExcelFile(source)
    else:
        xls = pd.ExcelFile(source, engine="xlrd")
    return xls.sheet_names


def _pick_sheet(source: BytesIO | Path, preferred: str) -> str:
    sheets = _list_sheets(source)
    if not sheets:
        raise ValueError("Spectra workbook has no sheets")
    if preferred in sheets:
        return preferred
    lower_map = {name.lower(): name for name in sheets}
    if preferred.lower() in lower_map:
        return lower_map[preferred.lower()]
    for name in sheets:
        if preferred.lower() in name.lower():
            return name
    return sheets[0]


def read_spectra_raw(source: BytesIO | Path) -> pd.DataFrame:
    sheet_name = _pick_sheet(source, DEFAULT_SHEET_NAME)
    return pd.read_excel(
        source,
        sheet_name=sheet_name,
        engine="xlrd",
        dtype=str,
        header=9,
    )


def _normalize_ticker_value(val: object) -> str:
    s = "" if val is None else str(val).strip().upper()
    if s in {"", "NAN"}:
        return ""
    parts = s.split()
    if len(parts) >= 2:
        base, market = parts[0], parts[1]
    else:
        base, market = s, ""
    market_map = {"UP": "US", "JT": "JP"}
    market = market_map.get(market, market)
    return (base + (" " + market if market else "")).strip()


def normalize_spectra(df: pd.DataFrame) -> pd.DataFrame:
    needed_cols = [
        "Shares/Par",
        "Price",
        "Traded Market Value",
        "Traded Market Value (Base)",
    ]
    type_col = df.columns[5] if len(df.columns) > 5 else df.columns[-1]
    id_col = df.columns[6] if len(df.columns) > 6 else df.columns[-1]

    keep_cols = [c for c in needed_cols if c in df.columns]
    work = df[[type_col, id_col] + keep_cols].copy()
    work.rename(columns={type_col: "_type_raw", id_col: "_id_raw"}, inplace=True)

    bond_set = SETTINGS.bond_code_set
    stack_set = SETTINGS.stack_code_set

    def classify_id_type(x: str) -> str:
        value = (x or "").strip()
        if value in bond_set:
            return "ISIN"
        if value in stack_set:
            return "TICKER"
        return "UNKNOWN"

    work["id_type"] = work["_type_raw"].map(classify_id_type)
    work["id_value"] = work["_id_raw"].astype(str).str.strip().str.upper()

    mask_unknown = ~work["id_type"].isin(["ISIN", "TICKER"])
    work = work[~mask_unknown].copy()

    mask_ticker = work["id_type"].str.upper() == "TICKER"
    work.loc[mask_ticker, "id_value"] = work.loc[mask_ticker, "id_value"].map(_normalize_ticker_value)
    return work


def spectra_to_records(
    source: BytesIO | Path,
    nav_date: date | None = None,
    currency: str | None = None,
    share_class: str = "DEFAULT",
) -> Sequence[NavRecord]:
    raw_bytes = ensure_bytes(source)
    dataframe = read_spectra_raw(BytesIO(raw_bytes))
    normalized = normalize_spectra(dataframe)
    nav_date = nav_date or extract_nav_date(normalized)
    currency = currency or extract_currency(normalized)
    share_class = share_class or extract_share_class(normalized)

    records: list[NavRecord] = []
    mapping = load_mapping()
    digest = compute_file_hash(raw_bytes)
    as_of = datetime.utcnow()

    for idx, row in normalized.iterrows():
        id_value = str(row.get("id_value", "")).strip().upper()
        nav_raw = row.get("Traded Market Value (Base)")
        nav_value = parse_decimal(nav_raw)
        security_id = mapping.get(id_value, id_value)
        if not security_id:
            continue
        identity = NavIdentity(
            instrument_id=security_id,
            nav_date=nav_date,
            currency=currency or "USD",
            share_class=share_class,
        )
        records.append(
            NavRecord(
                identity=identity,
                nav=nav_value,
                source="spectra",
                file_hash=digest,
                as_of=as_of,
                lineage=f"row={idx}",
            )
        )
    return records
