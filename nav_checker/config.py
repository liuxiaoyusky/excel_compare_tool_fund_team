"""Central configuration for the NAV checker package."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import timezone
from decimal import Context, Decimal
from pathlib import Path

from nav_checker.infrastructure.storage.mapping_store import load_mapping

# Legacy sets retained for compatibility; to be replaced with curated taxonomy.
BOND_CODE_SET = {"MR", "MF", "G", "CB", "B"}
STACK_CODE_SET = {"S", "P"}

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
RAW_DIR = DATA_DIR / "raw"
NORMALIZED_DIR = DATA_DIR / "normalized"
AUDIT_DIR = DATA_DIR / "audit"

for path in (DATA_DIR, RAW_DIR, NORMALIZED_DIR, AUDIT_DIR):
    path.mkdir(parents=True, exist_ok=True)


@dataclass(slots=True, frozen=True)
class Settings:
    decimal_context: Context
    tolerance_abs: Decimal
    timezone: timezone.__class__
    id_mapping: dict[str, str]
    bond_code_set: set[str]
    stack_code_set: set[str]


SETTINGS = Settings(
    decimal_context=Context(prec=28),
    tolerance_abs=Decimal("0"),
    timezone=timezone.utc,
    id_mapping=load_mapping(),
    bond_code_set=set(BOND_CODE_SET),
    stack_code_set=set(STACK_CODE_SET),
)
