from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from config import (
    bond_code_set,
    stack_code_set,
    missing_isin_or_stack_code_mapping_dict,
)


LEFT_SHEET_DEFAULT = "Security Distribution"
RIGHT_SHEET_DEFAULT = "HSBC Position Appraisal Report"

ColumnMapping = Dict[str, str]

DEFAULT_COLUMN_MAPPING: ColumnMapping = {
    # left -> right
    "Shares/Par": "Quantity",
    "Price": "Local Market Price",
    "Traded Market Value": "Local Market Value",
    "Traded Market Value (Base)": "Book Market Value",
}

LEFT_CANDIDATE_KEYS = [
    "Security SEDOL",
    "Security Number",
    "Security Description (Short)",
]

RIGHT_CANDIDATE_KEYS = [
    "Sedol",
    "Security ID",
    "Investment",
]


def _normalize_key_series(series: pd.Series) -> pd.Series:
    if series is None:
        return series
    s = series.astype(str)
    s = s.where(~series.isna(), other=np.nan)
    s = s.str.strip().str.upper()
    s = s.replace("", np.nan)
    return s


def _coerce_numeric(series: pd.Series) -> pd.Series:
    if series is None:
        return series
    if series.dtype == object:
        series = series.astype(str).str.replace(",", "", regex=False).str.strip()
        series = series.replace("", np.nan)
    return pd.to_numeric(series, errors="coerce")


@dataclass
class KeyCandidate:
    left_key: str
    right_key: str
    coverage_ratio: float
    left_unique_ratio: float
    right_unique_ratio: float
    left_non_null: int
    right_non_null: int
    matched_rows: int
    score: float


def _score_key_pair(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_key: str,
    right_key: str,
) -> Optional[KeyCandidate]:
    if left_key not in left_df.columns or right_key not in right_df.columns:
        return None

    lk = _normalize_key_series(left_df[left_key])
    rk = _normalize_key_series(right_df[right_key])

    left_non_null = int(lk.notna().sum())
    right_non_null = int(rk.notna().sum())
    if left_non_null == 0 or right_non_null == 0:
        return None

    left_unique_ratio = float(lk.dropna().nunique() / left_non_null)
    right_unique_ratio = float(rk.dropna().nunique() / right_non_null)

    left_keys = set(lk.dropna().unique().tolist())
    right_keys = set(rk.dropna().unique().tolist())
    intersection = left_keys & right_keys
    matched_rows = int(left_df[lk.isin(intersection)].shape[0])
    denom = max(1, min(left_non_null, right_non_null))
    coverage_ratio = float(matched_rows / denom)

    score = (
        coverage_ratio * 0.7
        + min(left_unique_ratio, right_unique_ratio) * 0.25
        + (1.0 if (left_unique_ratio > 0.98 and right_unique_ratio > 0.98) else 0.0) * 0.05
    )

    return KeyCandidate(
        left_key=left_key,
        right_key=right_key,
        coverage_ratio=coverage_ratio,
        left_unique_ratio=left_unique_ratio,
        right_unique_ratio=right_unique_ratio,
        left_non_null=left_non_null,
        right_non_null=right_non_null,
        matched_rows=matched_rows,
        score=score,
    )


def suggest_keys(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_candidates: Optional[List[str]] = None,
    right_candidates: Optional[List[str]] = None,
) -> List[KeyCandidate]:
    left_candidates = left_candidates or LEFT_CANDIDATE_KEYS
    right_candidates = right_candidates or RIGHT_CANDIDATE_KEYS

    results: List[KeyCandidate] = []
    for lk in left_candidates:
        for rk in right_candidates:
            kc = _score_key_pair(left_df, right_df, lk, rk)
            if kc is not None:
                results.append(kc)
    results.sort(key=lambda x: x.score, reverse=True)
    return results


def read_excel_sheet(file_like_or_path, sheet_name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_like_or_path, sheet_name=sheet_name, dtype=object)
    except ValueError as e:
        xls = pd.ExcelFile(file_like_or_path)
        raise ValueError(
            f"Sheet '{sheet_name}' 未找到。可用 sheets: {xls.sheet_names}"
        ) from e
    return df


def _prepare_numeric_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    exists = [c for c in columns if c in df.columns]
    for c in exists:
        df[c] = _coerce_numeric(df[c])
    return df


@dataclass
class CompareConfig:
    left_key: str
    right_key: str
    column_mapping: ColumnMapping


@dataclass
class CompareResult:
    diffs: pd.DataFrame
    only_in_left: pd.DataFrame
    only_in_right: pd.DataFrame
    summary: pd.DataFrame
    notes: Optional[str] = None


# --- 类型规则匹配：基于左F/G与右J/M/O ----------------------------------------------------

def _get_col_by_index(df: pd.DataFrame, idx: int) -> str:
    cols = list(df.columns)
    if idx >= len(cols):
        raise IndexError(f"列索引 {idx} 超出范围（共有 {len(cols)} 列）。")
    return cols[idx]


def compare_data_type_rules(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    column_mapping: ColumnMapping = DEFAULT_COLUMN_MAPPING,
) -> CompareResult:
    left_df = left_df.copy()
    right_df = right_df.copy()

    # 左侧：F(索引5)为类型，G(索引6)为标识
    left_type_col = _get_col_by_index(left_df, 5)
    left_id_col = _get_col_by_index(left_df, 6)

    # 右侧：J(索引9) Security ID，M(索引12) Isin，O(索引14) Ticker
    right_secid_col = _get_col_by_index(right_df, 9)
    right_isin_col = _get_col_by_index(right_df, 12)
    right_ticker_col = _get_col_by_index(right_df, 14)

    # 归一化
    left_df[left_type_col] = _normalize_key_series(left_df[left_type_col])
    left_df[left_id_col] = _normalize_key_series(left_df[left_id_col])
    right_df[right_secid_col] = _normalize_key_series(right_df[right_secid_col])
    right_df[right_isin_col] = _normalize_key_series(right_df[right_isin_col])
    right_df[right_ticker_col] = _normalize_key_series(right_df[right_ticker_col])

    # 基础索引（非空）
    isin_to_idx: Dict[str, int] = {}
    ticker_to_idx: Dict[str, int] = {}
    for i, val in right_df[right_isin_col].dropna().items():
        if val not in isin_to_idx:
            isin_to_idx[val] = i
    for i, val in right_df[right_ticker_col].dropna().items():
        if val not in ticker_to_idx:
            ticker_to_idx[val] = i

    # 缺失映射索引：仅同类型兜底
    bond_fallback_map: Dict[str, int] = {}
    stock_fallback_map: Dict[str, int] = {}
    for i, row in right_df.iterrows():
        secid = row.get(right_secid_col)
        if pd.isna(secid):
            continue
        mapped_val = missing_isin_or_stack_code_mapping_dict.get(secid)
        if not mapped_val:
            continue
        mapped_val = str(mapped_val).strip().upper()
        # 债券缺失：使用映射值作为候选 ISIN
        if pd.isna(row.get(right_isin_col)) and mapped_val:
            if mapped_val not in bond_fallback_map:
                bond_fallback_map[mapped_val] = i
        # 股票缺失：使用映射值作为候选 Ticker
        if pd.isna(row.get(right_ticker_col)) and mapped_val:
            if mapped_val not in stock_fallback_map:
                stock_fallback_map[mapped_val] = i

    left_type_series = left_df[left_type_col].fillna("")
    bond_codes = set(x.upper() for x in bond_code_set)
    stock_codes = set(x.upper() for x in stack_code_set)

    is_bond = left_type_series.isin(bond_codes)
    is_stock = left_type_series.isin(stock_codes)

    left_bonds = left_df[is_bond].copy()
    left_stocks = left_df[is_stock].copy()

    matched_records: List[Dict[str, object]] = []
    used_right_indices: set[int] = set()
    unmatched_rows: List[Dict[str, object]] = []

    def _append_record(lrow: pd.Series, rrow: pd.Series, matched_via: str):
        rec = {"key_left": lrow.get(left_id_col), "matched_via": matched_via}
        for lcol, rcol in column_mapping.items():
            rec[f"{lcol}__left"] = lrow.get(lcol)
            rec[f"{rcol}__right"] = rrow.get(rcol)
            lv = _coerce_numeric(pd.Series([lrow.get(lcol)])).iloc[0]
            rv = _coerce_numeric(pd.Series([rrow.get(rcol)])).iloc[0]
            if pd.isna(lv) and pd.isna(rv):
                diff = np.nan
                pct = np.nan
                equal = True
            else:
                if pd.notna(lv) and pd.notna(rv):
                    diff = lv - rv
                    pct = diff / rv if rv != 0 else np.nan
                    equal = float(lv) == float(rv)
                else:
                    diff = np.nan
                    pct = np.nan
                    equal = False
            rec[f"{lcol}__diff"] = diff
            rec[f"{lcol}__pct"] = pct
            rec[f"{lcol}__equal"] = bool(equal)
        matched_records.append(rec)

    # 债券：左G 对 右M(Isin)；若右M缺失，则按 Security ID 的映射值作为候选 Isin
    for _, lrow in left_bonds.iterrows():
        lid = lrow.get(left_id_col)
        if pd.isna(lid):
            continue
        # 直接按 ISIN
        ridx = isin_to_idx.get(lid)
        if ridx is not None:
            rrow = right_df.loc[ridx]
            used_right_indices.add(ridx)
            _append_record(lrow, rrow, matched_via="primary_isin")
            continue
        # 兜底：映射值作为 ISIN
        ridx2 = bond_fallback_map.get(str(lid))
        if ridx2 is not None:
            rrow2 = right_df.loc[ridx2]
            used_right_indices.add(ridx2)
            _append_record(lrow, rrow2, matched_via="fallback_isin_via_mapping")
            continue
        # 未匹配
        unmatched_rows.append({
            "left_type": "BOND",
            "left_id": lid,
            "reason": "no_match_or_missing_isin_without_mapping",
        })

    # 股票：左G 对 右O(Ticker)；若右O缺失，则按 Security ID 的映射值作为候选 Ticker
    for _, lrow in left_stocks.iterrows():
        lid = lrow.get(left_id_col)
        if pd.isna(lid):
            continue
        # 直接按 Ticker
        ridx = ticker_to_idx.get(lid)
        if ridx is not None:
            rrow = right_df.loc[ridx]
            used_right_indices.add(ridx)
            _append_record(lrow, rrow, matched_via="primary_ticker")
            continue
        # 兜底：映射值作为 Ticker
        ridx2 = stock_fallback_map.get(str(lid))
        if ridx2 is not None:
            rrow2 = right_df.loc[ridx2]
            used_right_indices.add(ridx2)
            _append_record(lrow, rrow2, matched_via="fallback_ticker_via_mapping")
            continue
        # 未匹配
        unmatched_rows.append({
            "left_type": "STOCK",
            "left_id": lid,
            "reason": "no_match_or_missing_ticker_without_mapping",
        })

    diffs_df = pd.DataFrame(matched_records)
    if not diffs_df.empty:
        diffs_df["any_diff"] = False
        for lcol in column_mapping.keys():
            eq_col = f"{lcol}__equal"
            if eq_col in diffs_df.columns:
                diffs_df["any_diff"] = diffs_df["any_diff"] | (~diffs_df[eq_col])
        diffs_df = diffs_df[diffs_df["any_diff"] == True]

    # 仅左/仅右/未匹配原因
    only_in_left = pd.DataFrame(unmatched_rows)
    only_in_right = right_df.loc[[i for i in range(len(right_df)) if i not in used_right_indices]].copy()

    summary_rows = []
    total_matches = int(len(matched_records))
    total_only_left = int(only_in_left.shape[0])
    total_only_right = int(only_in_right.shape[0])
    total_diffs = int(diffs_df.shape[0]) if not diffs_df.empty else 0
    summary_rows.append({"metric": "matched_rows", "value": total_matches})
    summary_rows.append({"metric": "only_in_left", "value": total_only_left})
    summary_rows.append({"metric": "only_in_right", "value": total_only_right})
    summary_rows.append({"metric": "rows_with_any_diff", "value": total_diffs})

    summary_df = pd.DataFrame(summary_rows)

    notes = (
        "类型规则：债券用左G对右M(Isin)；股票用左G对右O(Ticker)。"
        "当右侧对应标识为空时，仅在同类型内使用 Security ID 的映射值兜底；匹配失败即记为未匹配并报告。"
    )

    return CompareResult(
        diffs=diffs_df.reset_index(drop=True) if not diffs_df.empty else pd.DataFrame(),
        only_in_left=only_in_left.reset_index(drop=True),
        only_in_right=only_in_right.reset_index(drop=True),
        summary=summary_df,
        notes=notes,
    )


# --- 旧的通用键比对 -------------------------------------------------------------


def compare_data(left_df: pd.DataFrame, right_df: pd.DataFrame, config: CompareConfig) -> CompareResult:
    left_df = left_df.copy()
    right_df = right_df.copy()
    left_df[config.left_key] = _normalize_key_series(left_df[config.left_key])
    right_df[config.right_key] = _normalize_key_series(right_df[config.right_key])

    left_numeric_cols = list(config.column_mapping.keys())
    right_numeric_cols = list(config.column_mapping.values())
    left_df = _prepare_numeric_columns(left_df, left_numeric_cols)
    right_df = _prepare_numeric_columns(right_df, right_numeric_cols)

    left_keys = set(left_df[config.left_key].dropna().unique().tolist())
    right_keys = set(right_df[config.right_key].dropna().unique().tolist())

    only_left_keys = sorted(list(left_keys - right_keys))
    only_right_keys = sorted(list(right_keys - left_keys))

    only_in_left = left_df[left_df[config.left_key].isin(only_left_keys)].copy().sort_values(config.left_key)
    only_in_right = right_df[right_df[config.right_key].isin(only_right_keys)].copy().sort_values(config.right_key)

    merged = left_df.merge(
        right_df,
        left_on=config.left_key,
        right_on=config.right_key,
        how="inner",
        suffixes=("_left", "_right"),
    )

    diff_records = []
    for _, row in merged.iterrows():
        record = {
            "key_left": row.get(config.left_key),
            "key_right": row.get(config.right_key),
        }
        any_diff = False
        for lcol, rcol in config.column_mapping.items():
            lv = row.get(lcol, np.nan)
            rv = row.get(rcol, np.nan)
            if pd.isna(lv) and pd.isna(rv):
                diff = np.nan
                pct = np.nan
                equal = True
            else:
                if pd.notna(lv) and pd.notna(rv):
                    diff = lv - rv
                    pct = diff / rv if (rv != 0) else np.nan
                    try:
                        equal = float(lv) == float(rv)
                    except Exception:
                        equal = False
                else:
                    diff = np.nan
                    pct = np.nan
                    equal = False
            record[f"{lcol}__left"] = lv
            record[f"{rcol}__right"] = rv
            record[f"{lcol}__diff"] = diff
            record[f"{lcol}__pct"] = pct
            record[f"{lcol}__equal"] = bool(equal)
            if not equal:
                any_diff = True
        record["any_diff"] = any_diff
        diff_records.append(record)

    diffs_df = pd.DataFrame(diff_records)
    diffs_df = diffs_df[diffs_df["any_diff"] == True]

    summary_rows = []
    total_matches = int(merged.shape[0])
    total_only_left = int(only_in_left.shape[0])
    total_only_right = int(only_in_right.shape[0])
    total_diffs = int(diffs_df.shape[0])
    summary_rows.append({"metric": "matched_rows", "value": total_matches})
    summary_rows.append({"metric": "only_in_left", "value": total_only_left})
    summary_rows.append({"metric": "only_in_right", "value": total_only_right})
    summary_rows.append({"metric": "rows_with_any_diff", "value": total_diffs})
    for lcol, rcol in DEFAULT_COLUMN_MAPPING.items():
        merged_equal = int((
            _coerce_numeric(merged[lcol]).fillna(np.nan)
            == _coerce_numeric(merged[rcol]).fillna(np.nan)
        ).sum())
        summary_rows.append({"metric": f"{lcol}_equal_rows", "value": merged_equal})
        summary_rows.append({"metric": f"{lcol}_diff_rows", "value": int(total_matches - merged_equal)})
    summary_df = pd.DataFrame(summary_rows)

    return CompareResult(
        diffs=diffs_df.reset_index(drop=True),
        only_in_left=only_in_left.reset_index(drop=True),
        only_in_right=only_in_right.reset_index(drop=True),
        summary=summary_df,
    )


def export_to_excel_bytes(result: CompareResult) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.summary.to_excel(writer, sheet_name="summary", index=False)
        result.diffs.to_excel(writer, sheet_name="diffs", index=False)
        result.only_in_left.to_excel(writer, sheet_name="only_in_left", index=False)
        result.only_in_right.to_excel(writer, sheet_name="only_in_right", index=False)
    output.seek(0)
    return output.read()


# --- CLI entrypoint -----------------------------------------------------------
if __name__ == "__main__":
    import argparse
    import os
    import sys

    parser = argparse.ArgumentParser(description="Compare two Excel sheets with strict equality")
    parser.add_argument("--left", required=True, help="Left Excel file path")
    parser.add_argument("--right", required=True, help="Right Excel file path")
    parser.add_argument("--left-sheet", default=LEFT_SHEET_DEFAULT, help="Left sheet name")
    parser.add_argument("--right-sheet", default=RIGHT_SHEET_DEFAULT, help="Right sheet name")
    parser.add_argument("--left-key", default=None, help="Left join key column name (optional)")
    parser.add_argument("--right-key", default=None, help="Right join key column name (optional)")
    parser.add_argument("--output", default="compare_result.xlsx", help="Output Excel path")
    parser.add_argument("--join-mode", choices=["auto", "type_rules"], default="auto", help="Join strategy")

    args = parser.parse_args()

    if not os.path.exists(args.left):
        print(f"Left file not found: {args.left}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(args.right):
        print(f"Right file not found: {args.right}", file=sys.stderr)
        sys.exit(1)

    try:
        left_df = read_excel_sheet(args.left, sheet_name=args.left_sheet)
        right_df = read_excel_sheet(args.right, sheet_name=args.right_sheet)
    except Exception as e:
        print(f"读取 Excel 失败: {e}", file=sys.stderr)
        sys.exit(2)

    if args.join_mode == "type_rules":
        result = compare_data_type_rules(left_df, right_df, DEFAULT_COLUMN_MAPPING)
    else:
        left_key = args.left_key
        right_key = args.right_key
        if not left_key or not right_key:
            cands = suggest_keys(left_df, right_df)
            if not cands:
                print("无法自动推荐主键，请使用 --left-key 与 --right-key 指定，或使用 --join-mode type_rules。", file=sys.stderr)
                sys.exit(3)
            best = cands[0]
            left_key = best.left_key
            right_key = best.right_key
            print("自动选择主键:")
            print(f"  left_key = {left_key}")
            print(f"  right_key = {right_key}")
        config = CompareConfig(
            left_key=left_key,
            right_key=right_key,
            column_mapping=DEFAULT_COLUMN_MAPPING,
        )
        result = compare_data(left_df, right_df, config)

    print("\n=== Summary ===")
    try:
        print(result.summary.to_string(index=False))
        if result.notes:
            print("\nNotes:", result.notes)
    except Exception:
        pass

    try:
        data = export_to_excel_bytes(result)
        with open(args.output, "wb") as f:
            f.write(data)
        print(f"\n已生成对比结果: {args.output}")
    except Exception as e:
        print(f"导出失败: {e}", file=sys.stderr)
        sys.exit(5)