"""Shared parsing utilities for Excel ingestion."""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Iterable
import hashlib

import pandas as pd


def ensure_bytes(source: BytesIO | Path | bytes) -> bytes:
    if isinstance(source, bytes):
        return source
    if isinstance(source, BytesIO):
        return source.getvalue()
    if isinstance(source, Path):
        return source.read_bytes()
    raise TypeError(f"Unsupported source type: {type(source)!r}")


def compute_file_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def parse_decimal(value: object) -> Decimal:
    if value is None:
        return Decimal("0")
    s = str(value).strip()
    if not s:
        return Decimal("0")
    if s.upper() == "NAN":
        return Decimal("0")
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    for ch in [",", "$", "€", "£", " "]:
        s = s.replace(ch, "")
    try:
        result = Decimal(s)
    except InvalidOperation:
        return Decimal("0")
    if negative:
        result = -result
    return result.normalize()


KNOWN_DATE_COLUMNS = [
    "Valuation Date",
    "Value Date",
    "Date",
    "NAV Date",
]


def extract_nav_date(df: pd.DataFrame) -> date:
    for column in KNOWN_DATE_COLUMNS:
        if column in df.columns:
            series = pd.to_datetime(df[column], errors="coerce")
            first_valid = series.dropna().iloc[0] if not series.dropna().empty else None
            if first_valid is not None:
                return first_valid.date()
    return datetime.utcnow().date()


KNOWN_CURRENCY_COLUMNS = [
    "Base Currency",
    "Currency",
    "NAV Currency",
]


def extract_currency(df: pd.DataFrame) -> str:
    for column in KNOWN_CURRENCY_COLUMNS:
        if column in df.columns:
            series = df[column].dropna().astype(str).str.strip()
            if not series.empty:
                return series.iloc[0].upper()
    return "USD"


def extract_share_class(df: pd.DataFrame) -> str:
    for column in ("Share Class", "Class", "Class Name"):
        if column in df.columns:
            series = df[column].dropna().astype(str).str.strip()
            if not series.empty:
                return series.iloc[0]
    return "DEFAULT"
