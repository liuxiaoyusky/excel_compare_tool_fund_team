import sys
from pathlib import Path
from io import BytesIO

import pandas as pd
from decimal import Decimal, InvalidOperation

from config import (
    bond_code_set,
    stack_code_set,
    missing_isin_or_stack_code_mapping_dict,
    TOLERANCE_ABS,
)
import re


SPECTRA_PATH = Path("./spectra.xls")
SPECTRA_SHEET = "Security Distribution"

HSBC_PATH = Path("./HSBC Position Appraisal Report (EXCEL).xlsx")
HSBC_SHEET = "HSBC Position Appraisal Report"
def _normalize_ticker_value(val: object) -> str:
    s = "" if val is None else str(val).strip().upper()
    if s == "" or s == "NAN":
        return s
    # 去掉尾部 EQUITY
    s = re.sub(r"\s+EQUITY$", "", s)
    # 将连续空白规约为单空格
    s = re.sub(r"\s+", " ", s)
    parts = s.split(" ")
    if len(parts) >= 2:
        base, market = parts[0], parts[1]
    else:
        # 处理无空格的形态，如 7936JT
        m = re.match(r"^([0-9A-Z]+)([A-Z]{2})$", s)
        if m:
            base, market = m.group(1), m.group(2)
        else:
            base, market = s, ""
    # 市场后缀映射：UP->US, JT->JP
    market_map = {"UP": "US", "JT": "JP"}
    market = market_map.get(market, market)
    # 仅保留两段
    return (base + (" " + market if market else "")).strip()



def assert_files_exist() -> None:
    missing = []
    if not SPECTRA_PATH.exists():
        missing.append(str(SPECTRA_PATH))
    if not HSBC_PATH.exists():
        missing.append(str(HSBC_PATH))
    if missing:
        raise FileNotFoundError(f"Missing required files: {', '.join(missing)}")


def list_sheets(path: Path) -> list[str]:
    xls = pd.ExcelFile(path)
    return xls.sheet_names


def pick_sheet(path: Path, preferred: str) -> str:
    sheet_names = list_sheets(path)
    # 优先精确匹配
    if preferred in sheet_names:
        return preferred
    # 次选不区分大小写精确匹配
    lower_map = {name.lower(): name for name in sheet_names}
    if preferred.lower() in lower_map:
        return lower_map[preferred.lower()]
    # 再次选包含关系
    for name in sheet_names:
        if preferred.lower() in name.lower():
            return name
    # 最后退回第一个
    return sheet_names[0]


def read_spectra_raw() -> pd.DataFrame:
    chosen_sheet = pick_sheet(SPECTRA_PATH, SPECTRA_SHEET)
    # Header at row 10 (1-based), data starts row 11 → header=9, skiprows=range(0,9)
    df = pd.read_excel(
        SPECTRA_PATH,
        sheet_name=chosen_sheet,
        engine="xlrd",
        dtype=str,
        header=9,
    )
    return df


def read_hsbc_raw() -> pd.DataFrame:
    chosen_sheet = pick_sheet(HSBC_PATH, HSBC_SHEET)
    # Header at row 13 (1-based), data starts row 14 → header=12
    df = pd.read_excel(
        HSBC_PATH,
        sheet_name=chosen_sheet,
        engine="openpyxl",
        dtype=str,
        header=12,
    )
    return df


def spectra_normalize(df: pd.DataFrame) -> pd.DataFrame:
    # 明确使用 F 列作为类型，G 列作为 ID（按位置选择）
    needed_cols = [
        "Shares/Par",
        "Price",
        "Traded Market Value",
        "Traded Market Value (Base)",
    ]
    # F 列（索引 5）和 G 列（索引 6）
    type_col = df.columns[5] if len(df.columns) > 5 else df.columns[-1]
    id_col = df.columns[6] if len(df.columns) > 6 else df.columns[-1]

    # 过滤只保留需要列
    keep_cols = [c for c in needed_cols if c in df.columns]
    missing = [c for c in needed_cols if c not in df.columns]
    if missing:
        print(f"[spectra] 缺失必要列: {missing}")
    use_df = df[[type_col, id_col] + keep_cols].copy()
    use_df.rename(columns={type_col: "_type_raw", id_col: "_id_raw"}, inplace=True)

    # 生成统一键
    def classify_id_type(x: str) -> str:
        val = (x or "").strip()
        if val in bond_code_set:
            return "ISIN"
        if val in stack_code_set:
            return "TICKER"
        return "UNKNOWN"

    use_df["id_type"] = use_df["_type_raw"].map(classify_id_type)
    # ID 已在源提取阶段去掉币种，不做额外切分，只做标准化空白与大小写
    use_df["id_value"] = use_df["_id_raw"].astype(str).str.strip().str.upper()
    # 丢弃 UNKNOWN
    before = len(use_df)
    use_df = use_df[use_df["id_type"].isin(["ISIN", "TICKER"])].copy()
    dropped = before - len(use_df)
    if dropped:
        print(f"[spectra] 忽略 {dropped} 行（类型不在集合内）")

    # 只保留我们后续需要的列
    final_cols = [
        "id_type",
        "id_value",
        "_type_raw",
        "Shares/Par",
        "Price",
        "Traded Market Value",
        "Traded Market Value (Base)",
    ]
    final = use_df[[c for c in final_cols if c in use_df.columns]].copy()
    # 误分类纠正：若 id_type 为 ISIN 但值看起来是 ticker（如 SLV US），改判为 TICKER
    if "id_type" in final.columns and "id_value" in final.columns:
        val_series = final["id_value"].astype(str).str.strip().str.upper()
        looks_like_isin = val_series.str.match(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")
        looks_like_ticker = val_series.str.match(r"^[0-9A-Z]+\s+[A-Z]{2}$") | val_series.str.match(r"^[0-9A-Z]+(JP|JT|US|UP)$")
        mask_wrong_isin = (final["id_type"].astype(str).str.upper() == "ISIN") & (~looks_like_isin) & looks_like_ticker
        if mask_wrong_isin.any():
            final.loc[mask_wrong_isin, "id_type"] = "TICKER"
            final.loc[mask_wrong_isin, "id_value"] = final.loc[mask_wrong_isin, "id_value"].map(_normalize_ticker_value)
        # 反向纠正：若被标为 TICKER 但实际是 ISIN
        mask_wrong_ticker = (final["id_type"].astype(str).str.upper() == "TICKER") & looks_like_isin
        if mask_wrong_ticker.any():
            final.loc[mask_wrong_ticker, "id_type"] = "ISIN"
            # ISIN 保留空格前部分（若存在）
            final.loc[mask_wrong_ticker, "id_value"] = final.loc[mask_wrong_ticker, "id_value"].astype(str).str.split().str[0]
    # 规范化 spectra 侧的 id_value（针对 TICKER 类型做同样的市场后缀归一）
    if "id_type" in final.columns and "id_value" in final.columns:
        mask_ticker = final["id_type"].astype(str).str.upper() == "TICKER"
        final.loc[mask_ticker, "id_value"] = final.loc[mask_ticker, "id_value"].map(_normalize_ticker_value)
    return final


def main_step1_read_files() -> int:
    print("[Step 1] 读取配置与文件可用性检查…")
    # 仅访问配置，确保可导入
    _ = bond_code_set, stack_code_set, missing_isin_or_stack_code_mapping_dict
    assert_files_exist()

    print("[Step 1] 读取 spectra 原始数据…")
    spectra_sheets = list_sheets(SPECTRA_PATH)
    print("spectra sheets:", spectra_sheets)
    spectra_df = read_spectra_raw()
    print(f"spectra: shape={spectra_df.shape}")
    print("spectra columns (前20):", list(spectra_df.columns[:20]))
    spectra_norm = spectra_normalize(spectra_df)
    print(f"spectra normalized: shape={spectra_norm.shape}")
    print("spectra normalized columns:", list(spectra_norm.columns))

    print("[Step 1] 读取 HSBC 原始数据…")
    hsbc_sheets = list_sheets(HSBC_PATH)
    print("hsbc sheets:", hsbc_sheets)
    hsbc_df = read_hsbc_raw()
    print(f"hsbc: shape={hsbc_df.shape}")
    print("hsbc columns (前20):", list(hsbc_df.columns[:20]))

    print("[Step 1] OK")
    return 0


def hsbc_build_long_table(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # 缺失补全：基于 Security ID 映射，但配置仅给出一个 dict，无法区分 ISIN/Ticker。
    # 鉴于你的要求“补全明细不需要报告”，且我们仅用于连接，采取保守策略：
    # 若 Isin 为空且 Security ID 在映射中，填充值；若 Ticker 为空且 Security ID 在映射中，也填同一值。
    # 常见做法是分开两 dict，这里先按统一映射填充到两列，后续连接只用到其中之一。
    work = df.copy()
    for col in ["Isin", "Ticker"]:
        if col in work.columns:
            mask = work[col].isna() | (work[col].astype(str).str.strip() == "")
            work.loc[mask, col] = work.loc[mask, "Security ID"].map(
                missing_isin_or_stack_code_mapping_dict
            )

    # 规范化键值
    for col in ["Isin", "Ticker", "Security ID"]:
        if col in work.columns:
            work[col] = work[col].astype(str).str.strip().str.upper()

    # ISIN 可能是 "ISINNUM CURRENCY"，保留空格前部分
    if "Isin" in work.columns:
        work["Isin"] = work["Isin"].astype(str).str.split().str[0]

    # 统一规范化 HSBC 的 Ticker（EQUITY 去尾、两段裁剪、市场后缀映射等）
    if "Ticker" in work.columns:
        work["Ticker"] = work["Ticker"].map(_normalize_ticker_value)

    # 列映射确认（无容差比较所需的数值列）
    value_cols = [
        "Quantity",
        
        "Local Market Price",
        "Local Market Value",
        "Book Market Value",
    ]

    base_cols = [c for c in ["Security ID", "Isin", "Ticker"] if c in work.columns]
    keep_cols = base_cols + [c for c in value_cols if c in work.columns]
    base = work[keep_cols].copy()

    records: list[pd.DataFrame] = []
    if "Isin" in base.columns:
        mask_isin = base["Isin"].notna() & (base["Isin"].str.strip() != "") & (base["Isin"].str.upper() != "NAN")
        part_isin = base[mask_isin].copy()
        if not part_isin.empty:
            part_isin.insert(0, "id_type", "ISIN")
            part_isin.insert(1, "id_value", part_isin.pop("Isin"))
            records.append(part_isin)
    if "Ticker" in base.columns:
        mask_ticker = base["Ticker"].notna() & (base["Ticker"].str.strip() != "") & (base["Ticker"].str.upper() != "NAN")
        part_ticker = base[mask_ticker].copy()
        if not part_ticker.empty:
            part_ticker.insert(0, "id_type", "TICKER")
            part_ticker.insert(1, "id_value", part_ticker.pop("Ticker"))
            records.append(part_ticker)

    if records:
        long_df = pd.concat(records, ignore_index=True, sort=False)
    else:
        long_df = pd.DataFrame(columns=["id_type", "id_value"] + keep_cols)

    # 检测重复键
    dup_mask = long_df.duplicated(subset=["id_type", "id_value"], keep=False)
    dups = long_df.loc[dup_mask].copy()
    return long_df, dups, base


def left_join_non_dup(
    spectra_norm: pd.DataFrame,
    hsbc_long: pd.DataFrame,
    hsbc_dups: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # 排除重复键
    if not hsbc_dups.empty:
        dup_keys = hsbc_dups[["id_type", "id_value"]].drop_duplicates()
        hsbc_long = hsbc_long.merge(dup_keys, on=["id_type", "id_value"], how="left", indicator=True)
        hsbc_long = hsbc_long[hsbc_long["_merge"] == "left_only"].drop(columns=["_merge"])  # anti-join

    merged = spectra_norm.merge(
        hsbc_long,
        on=["id_type", "id_value"],
        how="left",
        suffixes=("_spectra", "_hsbc"),
    )
    # 未匹配：HSBC 四列值均为空（或缺列时以 Security ID 是否空为判定）
    hsbc_value_cols = [
        col for col in [
            "Quantity",
            "Local Market Price",
            "Local Market Value",
            "Book Market Value",
            "Security ID",
        ] if col in merged.columns
    ]
    if hsbc_value_cols:
        unmatched_mask = merged[hsbc_value_cols].isna().all(axis=1)
    else:
        unmatched_mask = pd.Series([True] * len(merged), index=merged.index)
    unmatched = merged[unmatched_mask].copy()
    return merged, unmatched


def apply_fallback_mapping(
    merged: pd.DataFrame,
    hsbc_base: pd.DataFrame,
) -> tuple[pd.DataFrame, int, int]:
    # 对于未匹配行：用 id_value 在 mapping 中查 Security ID；
    # 若命中，则在 hsbc_base 以 Security ID 精确查找；
    # - 若找到 1 行：回填右侧列（Quantity/Price/Values），不改变输出结构；
    # - 若找到多行：视为重复，跳过该条（不匹配）。
    work = merged.copy()
    updated = 0
    dup_skipped = 0
    # 构建 Security ID 索引映射到行位置（可能一对多）
    if "Security ID" not in hsbc_base.columns:
        return work, 0, 0
    secid_groups = hsbc_base.groupby("Security ID").indices

    hsbc_cols_to_copy = [c for c in ["Security ID", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in hsbc_base.columns]
    numeric_cols_to_copy = [c for c in ["Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in hsbc_base.columns]

    for idx, row in work.iterrows():
        # 若四个数值列均已有值，则跳过；仅 Security ID 有值但数值列缺失时仍允许回填
        if numeric_cols_to_copy:
            numeric_filled_all = all(
                (pd.notna(row.get(c)) and str(row.get(c)).strip() != "") for c in numeric_cols_to_copy
            )
            if numeric_filled_all:
                continue
        key_in = str(row["id_value"]).strip().upper()
        mapped_secid = missing_isin_or_stack_code_mapping_dict.get(key_in)
        if not mapped_secid:
            continue
        mapped_secid_up = str(mapped_secid).strip().upper()
        if mapped_secid_up not in secid_groups:
            continue
        positions = secid_groups[mapped_secid_up]
        if isinstance(positions, list) or hasattr(positions, "__len__") and len(positions) > 1:
            dup_skipped += 1
            continue
        # 单行匹配
        if hasattr(positions, "__len__") and len(positions) == 1:
            pos = positions[0]
        else:
            # 某些 pandas 版本 groupby.indices 返回 ndarray
            try:
                pos = positions[0]
            except Exception:
                continue
        rhs = hsbc_base.iloc[pos]
        for c in hsbc_cols_to_copy:
            work.at[idx, c] = rhs.get(c)
        updated += 1

    return work, updated, dup_skipped


def _to_decimal(value: str) -> Decimal | None:
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.upper() == "NAN":
        return None
    # 处理会计负号与千分位
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    # 清理符号
    for ch in [",", "$", "€", "£", "%", " "]:
        s = s.replace(ch, "")
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    if negative:
        d = -d
    # 规范化，去除多余 0
    return d.normalize()


SPECTRA_TO_HSBC_MAP = {
    "Shares/Par": "Quantity",
    "Price": "Local Market Price",
    "Traded Market Value": "Local Market Value",
    "Traded Market Value (Base)": "Book Market Value",
}

# 供三源对比使用的 VPFS 列映射（在三源合并前会将 VPFS 列重命名为以下列名）
SPECTRA_TO_VPFS_MAP = {
    "Shares/Par": "Quantity (VPFS)",
    "Price": "Local Market Price (VPFS)",
    "Traded Market Value": "Local Market Value (VPFS)",
    "Traded Market Value (Base)": "Book Market Value (VPFS)",
}


def build_comparison(merged: pd.DataFrame) -> pd.DataFrame:
    df = merged.copy()
    try:
        abs_tol = Decimal(str(TOLERANCE_ABS))
    except Exception:
        abs_tol = Decimal("0.001")
    for spectra_col, hsbc_col in SPECTRA_TO_HSBC_MAP.items():
        lhs_col = f"{spectra_col}"
        rhs_col = f"{hsbc_col}"
        if lhs_col not in df.columns:
            df[lhs_col] = None
        if rhs_col not in df.columns:
            df[rhs_col] = None

        # 原始源值列保留
        # 解析为 Decimal 并做严格等价比较
        lhs_num = df[lhs_col].map(_to_decimal)
        rhs_num = df[rhs_col].map(_to_decimal)
        # 等价：二者皆为 None → True；二者皆数值且 |b-a| <= 0.001 → True；其它 → False
        equal_flags: list[bool] = []
        for a, b in zip(lhs_num, rhs_num):
            if a is None and b is None:
                equal_flags.append(True)
            elif a is None or b is None:
                equal_flags.append(False)
            else:
                equal_flags.append(abs(b - a) <= abs_tol)
        df[f"{spectra_col}__equal"] = equal_flags
        # 差值（hsbc - spectra），若任一侧为空则为空
        deltas: list[Decimal | None] = []
        for a, b in zip(lhs_num, rhs_num):
            if a is None or b is None:
                deltas.append(None)
            else:
                deltas.append((b - a).normalize())
        df[f"{spectra_col}__delta"] = deltas

    # 是否存在任一列差异
    eq_cols = [f"{c}__equal" for c in SPECTRA_TO_HSBC_MAP.keys()]
    df["has_diff"] = ~df[eq_cols].all(axis=1)
    return df


def build_adjacent_export_df(comp: pd.DataFrame) -> pd.DataFrame:
    # 构造“成对相邻”的导出 DataFrame
    # 补充导出所需的 ID 列：
    # - spectra security ID 来自左侧规范化后的 id_value（若无则填 None）
    # - HSBC security ID 显示为右侧的 Security ID（保留内部列名不变，仅导出时映射）
    if "spectra security ID" not in comp.columns:
        comp = comp.assign(**{
            "spectra security ID": (comp["id_value"] if "id_value" in comp.columns else pd.Series(dtype=object))
        })
    if "HSBC security ID" not in comp.columns and "Security ID" in comp.columns:
        comp = comp.assign(**{
            "HSBC security ID": comp["Security ID"]
        })

    front_cols = [c for c in [
        "id_type", "id_value", "_type_raw",
        "spectra security ID", "HSBC security ID", "VPFS_security ID", "Isin", "Ticker"
    ] if c in comp.columns]
    ordered_cols: list[str] = front_cols.copy()

    for spectra_col, hsbc_col in SPECTRA_TO_HSBC_MAP.items():
        # 创建别名列，确保存在
        left_alias = f"{spectra_col}__spectra"
        right_alias = f"{spectra_col}__hsbc"
        delta_col = f"{spectra_col}__delta"
        equal_col = f"{spectra_col}__equal"

        if spectra_col not in comp.columns:
            comp.loc[:, spectra_col] = pd.Series(dtype=object)
        if hsbc_col not in comp.columns:
            comp.loc[:, hsbc_col] = pd.Series(dtype=object)

        comp[left_alias] = comp[spectra_col]
        comp[right_alias] = comp[hsbc_col]

        # 追加列顺序：spectra 值、hsbc 值、delta、equal
        ordered_cols += [left_alias, right_alias]
        if delta_col in comp.columns:
            ordered_cols.append(delta_col)
        if equal_col in comp.columns:
            ordered_cols.append(equal_col)

    # 在末尾加 has_diff，便于筛选
    if "has_diff" in comp.columns:
        ordered_cols.append("has_diff")

    # 去重并仅选择存在的列
    seen = set()
    final_cols = []
    for c in ordered_cols:
        if c in comp.columns and c not in seen:
            seen.add(c)
            final_cols.append(c)

    return comp[final_cols].copy()


def build_diff_columns_series(comp: pd.DataFrame) -> pd.Series:
    # 返回每行逗号分隔的不同 spectra 列名列表；若为未匹配行，则输出“{id_value}不存在”
    df = comp.copy()
    names = list(SPECTRA_TO_HSBC_MAP.keys())
    equal_cols = [f"{n}__equal" for n in names]
    # 若缺列则先补
    for col in equal_cols:
        if col not in df.columns:
            df.loc[:, col] = False

    # 识别未匹配行：HSBC 四列全空
    hsbc_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in df.columns]
    if hsbc_cols:
        unmatched_mask = df[hsbc_cols].isna().all(axis=1)
    else:
        unmatched_mask = pd.Series([False] * len(df), index=df.index)

    def row_to_list(r):
        diffs = []
        for n in names:
            col = f"{n}__equal"
            val = r.get(col)
            is_equal = False
            try:
                if pd.notna(val):
                    is_equal = bool(val)
            except Exception:
                is_equal = False
            if not is_equal:
                diffs.append(n)
        return ", ".join(diffs)

    res = df.apply(row_to_list, axis=1)
    if unmatched_mask.any():
        res.loc[unmatched_mask] = (
            df.loc[unmatched_mask, "id_value"].astype(str).str.strip() + "不存在"
        )
    return res


def _col_idx_to_excel(col_idx: int) -> str:
    # 0 -> A, 25 -> Z, 26 -> AA
    name = ""
    n = col_idx
    while True:
        n, r = divmod(n, 26)
        name = chr(65 + r) + name
        if n == 0:
            break
        n -= 1
    return name


def write_diffs_excel(export_diffs: pd.DataFrame, out_path: Path | BytesIO) -> None:
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        sheet = "diffs"
        (export_diffs if export_diffs is not None else pd.DataFrame()).to_excel(writer, sheet_name=sheet, index=False)
        if export_diffs is None or export_diffs.empty:
            return
        workbook  = writer.book
        worksheet = writer.sheets[sheet]
        yellow = workbook.add_format({"bg_color": "#FFFF00"})

        rows = len(export_diffs)
        # 对每个字段，若 equal 列为 False，则为该行对应的 spectra/hsbc 值列着色（逐行避免偏移）
        for spectra_col in SPECTRA_TO_HSBC_MAP.keys():
            left_alias = f"{spectra_col}__spectra"
            right_alias = f"{spectra_col}__hsbc"
            equal_col = f"{spectra_col}__equal"
            if left_alias not in export_diffs.columns or right_alias not in export_diffs.columns or equal_col not in export_diffs.columns:
                continue

            left_idx = export_diffs.columns.get_loc(left_alias)
            right_idx = export_diffs.columns.get_loc(right_alias)
            equal_idx = export_diffs.columns.get_loc(equal_col)
            equal_letter = _col_idx_to_excel(equal_idx)

            for r in range(rows):
                excel_row = r + 2  # 数据第1行对应 Excel 第2行
                formula = f"=${equal_letter}{excel_row}=FALSE"
                row_idx = r + 1  # xlsxwriter 的 0-based 行号（跳过表头）
                worksheet.conditional_format(row_idx, left_idx, row_idx, left_idx, {
                    "type": "formula",
                    "criteria": formula,
                    "format": yellow,
                })
                worksheet.conditional_format(row_idx, right_idx, row_idx, right_idx, {
                    "type": "formula",
                    "criteria": formula,
                    "format": yellow,
                })


def build_single_sheet_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    buf.seek(0)
    return buf.getvalue()


def _list_sheets_from_source(source: Path | BytesIO) -> list[str]:
    # 智能引擎选择：先尝试 openpyxl，失败则回退到 xlrd
    engines_to_tlry = ["openpyxl", "xlrd"]
    last_error = None
    
    for engine in engines_to_tlry:
        try:
            xls = pd.ExcelFile(source, engine=engine)
            return xls.sheet_names
        except Exception as e:
            last_error = e
            continue
    
    # 如果所有引擎都失败，抛出最后一个错误
    raise Exception(f"无法读取 Excel 文件工作表列表，尝试了所有引擎。最后错误: {last_error}")


def _pick_sheet_from_source(source: Path | BytesIO, preferred: str) -> str:
    sheet_names = _list_sheets_from_source(source)
    if preferred in sheet_names:
        return preferred
    lower_map = {name.lower(): name for name in sheet_names}
    if preferred.lower() in lower_map:
        return lower_map[preferred.lower()]
    for name in sheet_names:
        if preferred.lower() in name.lower():
            return name
    return sheet_names[0]


def read_spectra_raw_from(source: Path | BytesIO) -> pd.DataFrame:
    chosen_sheet = _pick_sheet_from_source(source, SPECTRA_SHEET)
    # 智能引擎选择：先尝试 openpyxl，失败则回退到 xlrd
    engines_to_tlry = ["openpyxl", "xlrd"]
    last_error = None
    
    for engine in engines_to_tlry:
        try:
            return pd.read_excel(source, sheet_name=chosen_sheet, engine=engine, dtype=str, header=9)
        except Exception as e:
            last_error = e
            continue
    
    # 如果所有引擎都失败，抛出最后一个错误
    raise Exception(f"无法读取 Excel 文件，尝试了所有引擎。最后错误: {last_error}")


def read_hsbc_raw_from(source: Path | BytesIO) -> pd.DataFrame:
    chosen_sheet = _pick_sheet_from_source(source, HSBC_SHEET)
    # 智能引擎选择：先尝试 openpyxl，失败则回退到 xlrd
    engines_to_tlry = ["openpyxl", "xlrd"]
    last_error = None
    
    for engine in engines_to_tlry:
        try:
            return pd.read_excel(source, sheet_name=chosen_sheet, engine=engine, dtype=str, header=12)
        except Exception as e:
            last_error = e
            continue
    
    # 如果所有引擎都失败，抛出最后一个错误
    raise Exception(f"无法读取 Excel 文件，尝试了所有引擎。最后错误: {last_error}")


# === VPFS 支持（d01 2.xls） ===
def read_vpfs_raw_from(source: Path | BytesIO) -> pd.DataFrame:
    # VPFS 文件通常为 .xls，采用第一个工作表
    # 行头未知，按默认首行作为表头读取，保留字符串类型
    chosen_sheet = _pick_sheet_from_source(source, "")
    # 智能引擎选择：先尝试 xlrd（适合 .xls），失败则尝试 openpyxl
    engines_to_tlry = ["xlrd", "openpyxl"]
    last_error = None
    
    for engine in engines_to_tlry:
        try:
            return pd.read_excel(source, sheet_name=chosen_sheet, engine=engine, dtype=str)
        except Exception as e:
            last_error = e
            continue
    
    # 如果所有引擎都失败，抛出最后一个错误
    raise Exception(f"无法读取 VPFS Excel 文件，尝试了所有引擎。最后错误: {last_error}")


def vpfs_normalize(df: pd.DataFrame) -> pd.DataFrame:
    # 列名归一化为小写并去空白，避免大小写与空格差异
    cols_map = {c: str(c).strip() for c in df.columns}
    df = df.rename(columns=cols_map)
    lower_map = {str(c).strip(): str(c).strip().lower() for c in df.columns}
    df.columns = [lower_map[str(c).strip()] for c in df.columns]

    work = df.copy()
    # 过滤 rowid == listing（不区分大小写）
    if "rowid" in work.columns:
        work = work[work["rowid"].astype(str).str.strip().str.lower() == "listing"].copy()
    # 丢弃 prodid 为空
    if "prodid" in work.columns:
        work = work[work["prodid"].astype(str).str.strip() != ""].copy()
        # 跳过 prodid == zzzzzz（不区分大小写）
        mask_skip = work["prodid"].astype(str).str.strip().str.lower() == "zzzzzz"
        if mask_skip.any():
            work = work[~mask_skip].copy()
    else:
        # 无 prodid 列则返回空框架，后续合并为完全未匹配
        return pd.DataFrame(columns=["prodid_norm", "VPFS_security ID", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"])  # type: ignore[list-item]

    # 规范化 prodid 作为连接键
    work.loc[:, "prodid_norm"] = work["prodid"].astype(str).str.strip().str.upper()
    # 添加 VPFS_security ID 列，值为原始的 prodid
    work.loc[:, "VPFS_security ID"] = work["prodid"]

    # 将 VPFS 数值列映射为右侧标准列（沿用 HSBC 列名，以便复用比较/导出逻辑）
    # 映射关系：valqty→Quantity，valprice→Local Market Price，MarketValue→Local Market Value，FundCCYMarketValue→Book Market Value
    target_cols = {
        "valqty": "Quantity",
        "valprice": "Local Market Price",
        "marketvalue": "Local Market Value",
        "fundccymarketvalue": "Book Market Value",
    }
    present = {src: tgt for src, tgt in target_cols.items() if src in work.columns}
    use_cols = ["prodid_norm", "VPFS_security ID"] + list(present.keys())
    use = work[use_cols].copy()
    use = use.rename(columns=present)
    # 确保所有目标列存在（缺失列补空）
    for tgt in target_cols.values():
        if tgt not in use.columns:
            use.loc[:, tgt] = pd.Series(dtype=object)
    
    return use[["prodid_norm", "VPFS_security ID", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"]].copy()


def run_compare_spectra_vs_vpfs(spectra_src: Path | BytesIO, vpfs_src: Path | BytesIO) -> dict:
    # 读取并标准化左侧 Spectra
    spectra_df = read_spectra_raw_from(spectra_src)
    spectra_norm = spectra_normalize(spectra_df)
    # 读取并标准化 VPFS（右侧）
    vpfs_df = read_vpfs_raw_from(vpfs_src)
    vpfs_norm = vpfs_normalize(vpfs_df)

    # 以 id_value（左）对 prodid_norm（右）左连接
    merged = spectra_norm.merge(
        vpfs_norm,
        left_on="id_value",
        right_on="prodid_norm",
        how="left",
        suffixes=("_spectra", "_vpfs"),
    )
    # 比较（沿用 SPECTRA_TO_HSBC_MAP）
    comp = build_comparison(merged)
    export_comp = build_adjacent_export_df(comp)
    export_comp.loc[:, "diff_columns"] = build_diff_columns_series(comp)
    export_diffs = export_comp[export_comp.get("has_diff", pd.Series([False] * len(export_comp))).astype(bool)].copy()

    # 左侧未匹配：右侧四列均为空
    rhs_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in merged.columns]
    if rhs_cols:
        unmatched_mask_export = merged[rhs_cols].isna().all(axis=1)
    else:
        unmatched_mask_export = pd.Series([True] * len(merged), index=merged.index)
    left_unmatched = merged[unmatched_mask_export].copy()
    export_unmatched_left = build_adjacent_export_df(build_comparison(left_unmatched.copy()) if not left_unmatched.empty else left_unmatched.copy())
    if not export_unmatched_left.empty:
        export_unmatched_left.loc[:, "diff_columns"] = build_diff_columns_series(build_comparison(left_unmatched.copy()))
        export_unmatched_left.loc[:, "source"] = "left"
    else:
        export_unmatched_left = export_unmatched_left.assign(**{
            "diff_columns": pd.Series(dtype=object),
            "source": pd.Series(dtype=object),
        })

    # 右侧未覆盖：VPFS prodid_norm 不在 spectra 的 id_value 中
    spectra_ids = set(export_comp.get("spectra security ID", pd.Series(dtype=str)).astype(str).str.strip().str.upper().dropna().tolist())
    if "prodid_norm" in vpfs_norm.columns:
        vpfs_only = vpfs_norm[~vpfs_norm["prodid_norm"].astype(str).str.strip().str.upper().isin(spectra_ids)].copy()
    else:
        vpfs_only = vpfs_norm.iloc[0:0].copy()
    comp_rhs_only = build_comparison(vpfs_only.copy()) if not vpfs_only.empty else vpfs_only.copy()
    export_unmatched_right = build_adjacent_export_df(comp_rhs_only) if not comp_rhs_only.empty else comp_rhs_only.copy()
    if not export_unmatched_right.empty:
        export_unmatched_right.loc[:, "diff_columns"] = build_diff_columns_series(comp_rhs_only)
        export_unmatched_right.loc[:, "source"] = "right"
    else:
        export_unmatched_right = export_unmatched_right.assign(**{
            "diff_columns": pd.Series(dtype=object),
            "source": pd.Series(dtype=object),
        })

    export_unmatched = pd.concat([export_unmatched_left, export_unmatched_right], ignore_index=True, sort=False)
    # VPFS 无重复键检测，提供空表结构
    export_dups = pd.DataFrame(columns=export_comp.columns)

    all_bytes = build_all_sheets_bytes(export_comp, export_diffs, export_unmatched, export_dups)
    return {
        "comparison": export_comp,
        "diffs": export_diffs,
        "unmatched": export_unmatched,
        "duplicates": export_dups,
        "all_sheets_xlsx": all_bytes,
    }


def build_all_sheets_bytes(
    export_comp: pd.DataFrame,
    export_diffs: pd.DataFrame,
    export_unmatched: pd.DataFrame,
    export_dups: pd.DataFrame,
) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        export_comp.to_excel(writer, sheet_name="comparison", index=False)
        export_unmatched.to_excel(writer, sheet_name="unmatched", index=False)
        export_dups.to_excel(writer, sheet_name="duplicates", index=False)
        # diffs with highlight
        sheet = "diffs"
        export_diffs.to_excel(writer, sheet_name=sheet, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet]
        yellow = workbook.add_format({"bg_color": "#FFFF00"})
        rows = len(export_diffs)
        for spectra_col in SPECTRA_TO_HSBC_MAP.keys():
            left_alias = f"{spectra_col}__spectra"
            right_alias = f"{spectra_col}__hsbc"
            equal_col = f"{spectra_col}__equal"
            if left_alias not in export_diffs.columns or right_alias not in export_diffs.columns or equal_col not in export_diffs.columns:
                continue
            left_idx = export_diffs.columns.get_loc(left_alias)
            right_idx = export_diffs.columns.get_loc(right_alias)
            equal_idx = export_diffs.columns.get_loc(equal_col)
            equal_letter = _col_idx_to_excel(equal_idx)
            for r in range(rows):
                excel_row = r + 2
                formula = f"=${equal_letter}{excel_row}=FALSE"
                row_idx = r + 1
                worksheet.conditional_format(row_idx, left_idx, row_idx, left_idx, {
                    "type": "formula",
                    "criteria": formula,
                    "format": yellow,
                })
                worksheet.conditional_format(row_idx, right_idx, row_idx, right_idx, {
                    "type": "formula",
                    "criteria": formula,
                    "format": yellow,
                })
    buf.seek(0)
    return buf.getvalue()


def run_compare_from_sources(spectra_src: Path | BytesIO, hsbc_src: Path | BytesIO) -> dict:
    spectra_df = read_spectra_raw_from(spectra_src)
    spectra_norm = spectra_normalize(spectra_df)
    hsbc_df = read_hsbc_raw_from(hsbc_src)
    long_df, dups, hsbc_base = hsbc_build_long_table(hsbc_df)
    merged, _ = left_join_non_dup(spectra_norm, long_df, dups)
    merged_fallback, _, _ = apply_fallback_mapping(merged, hsbc_base)
    comp = build_comparison(merged_fallback)
    export_comp = build_adjacent_export_df(comp)
    export_comp.loc[:, "diff_columns"] = build_diff_columns_series(comp)
    export_diffs = export_comp[export_comp["has_diff"]].copy() if "has_diff" in export_comp.columns else export_comp.iloc[0:0].copy()
    # unmatched：左侧未匹配 + 右侧未覆盖（基于 HSBC security ID 未出现于 comparison）
    hsbc_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in merged_fallback.columns]
    if hsbc_cols:
        unmatched_mask_export = merged_fallback[hsbc_cols].isna().all(axis=1)
    else:
        unmatched_mask_export = pd.Series([True] * len(merged_fallback), index=merged_fallback.index)
    left_unmatched = merged_fallback[unmatched_mask_export].copy()
    export_unmatched_left = build_adjacent_export_df(left_unmatched)
    if not export_unmatched_left.empty:
        export_unmatched_left.loc[:, "diff_columns"] = build_diff_columns_series(build_comparison(left_unmatched.copy()))
        export_unmatched_left.loc[:, "source"] = "left"
    else:
        # 保持列结构，避免后续 concat 类型不一致
        export_unmatched_left = export_unmatched_left.assign(**{
            "diff_columns": pd.Series(dtype=object),
            "source": pd.Series(dtype=object),
        })

    # 右侧未覆盖：比较 hsbc_base 的 Security ID 与 comparison 的 HSBC security ID
    comp_hsbc_ids = set(export_comp.get("HSBC security ID", pd.Series(dtype=str)).astype(str).str.strip().str.upper().dropna().tolist())
    if "Security ID" in hsbc_base.columns:
        hsbc_ids_all = hsbc_base["Security ID"].astype(str).str.strip().str.upper()
        rhs_missing = hsbc_base[~hsbc_ids_all.isin(comp_hsbc_ids)].copy()
    else:
        rhs_missing = hsbc_base.iloc[0:0].copy()
    comp_rhs_only = build_comparison(rhs_missing.copy()) if not rhs_missing.empty else rhs_missing.copy()
    export_unmatched_right = build_adjacent_export_df(comp_rhs_only) if not comp_rhs_only.empty else comp_rhs_only.copy()
    if not export_unmatched_right.empty:
        export_unmatched_right.loc[:, "diff_columns"] = build_diff_columns_series(comp_rhs_only)
        export_unmatched_right.loc[:, "source"] = "right"
    else:
        export_unmatched_right = export_unmatched_right.assign(**{
            "diff_columns": pd.Series(dtype=object),
            "source": pd.Series(dtype=object),
        })

    export_unmatched = pd.concat([export_unmatched_left, export_unmatched_right], ignore_index=True, sort=False)
    dups_for_comp = dups.copy()
    dups_for_comp = dups_for_comp[[c for c in ["id_type", "id_value", "Security ID", "Isin", "Ticker", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in dups_for_comp.columns]]
    dups_comp = build_comparison(dups_for_comp)
    export_dups = build_adjacent_export_df(dups_comp)
    export_dups.loc[:, "diff_columns"] = build_diff_columns_series(dups_comp)

    all_bytes = build_all_sheets_bytes(export_comp, export_diffs, export_unmatched, export_dups)
    return {
        "comparison": export_comp,
        "diffs": export_diffs,
        "unmatched": export_unmatched,
        "duplicates": export_dups,
        "all_sheets_xlsx": all_bytes,
    }


# ===== 三源对比：Spectra ↔ HSBC ↔ VPFS =====
def _vpfs_for_triple(vpfs_norm: pd.DataFrame) -> pd.DataFrame:
    # 将 VPFS 标准列重命名为带 (VPFS) 后缀，避免与 HSBC 列冲突
    rename_map = {
        "Quantity": "Quantity (VPFS)",
        "Local Market Price": "Local Market Price (VPFS)",
        "Local Market Value": "Local Market Value (VPFS)",
        "Book Market Value": "Book Market Value (VPFS)",
    }
    work = vpfs_norm.copy()
    for src, tgt in rename_map.items():
        if src in work.columns:
            work = work.rename(columns={src: tgt})
        else:
            work.loc[:, tgt] = pd.Series(dtype=object)
    
    # 保留 VPFS_security ID 列（如果存在）
    if "VPFS_security ID" not in work.columns:
        work["VPFS_security ID"] = pd.Series(dtype=object)
    
    return work


def build_triple_comparison(merged: pd.DataFrame) -> pd.DataFrame:
    df = merged.copy()
    try:
        abs_tol = Decimal(str(TOLERANCE_ABS))
    except Exception:
        abs_tol = Decimal("0.001")
    # 针对每个字段，对 HSBC 与 VPFS 分别计算等价与差值
    for spectra_col in SPECTRA_TO_HSBC_MAP.keys():
        lhs_col = spectra_col
        hsbc_col = SPECTRA_TO_HSBC_MAP[spectra_col]
        vpfs_col = SPECTRA_TO_VPFS_MAP[spectra_col]
        if lhs_col not in df.columns:
            df[lhs_col] = None
        if hsbc_col not in df.columns:
            df[hsbc_col] = None
        if vpfs_col not in df.columns:
            df[vpfs_col] = None

        lhs_num = df[lhs_col].map(_to_decimal)
        rhs_hsbc = df[hsbc_col].map(_to_decimal)
        rhs_vpfs = df[vpfs_col].map(_to_decimal)

        eq_hsbc: list[bool] = []
        eq_vpfs: list[bool] = []
        delta_hsbc: list[Decimal | None] = []
        delta_vpfs: list[Decimal | None] = []
        for a, b_h, b_v in zip(lhs_num, rhs_hsbc, rhs_vpfs):
            # hsbc
            if a is None and b_h is None:
                eq_hsbc.append(True)
            elif a is None or b_h is None:
                eq_hsbc.append(False)
            else:
                eq_hsbc.append(abs(b_h - a) <= abs_tol)
            if a is None or b_h is None:
                delta_hsbc.append(None)
            else:
                delta_hsbc.append((b_h - a).normalize())
            # vpfs
            if a is None and b_v is None:
                eq_vpfs.append(True)
            elif a is None or b_v is None:
                eq_vpfs.append(False)
            else:
                eq_vpfs.append(abs(b_v - a) <= abs_tol)
            if a is None or b_v is None:
                delta_vpfs.append(None)
            else:
                delta_vpfs.append((b_v - a).normalize())

        df[f"{spectra_col}__equal_hsbc"] = eq_hsbc
        df[f"{spectra_col}__delta_hsbc"] = delta_hsbc
        df[f"{spectra_col}__equal_vpfs"] = eq_vpfs
        df[f"{spectra_col}__delta_vpfs"] = delta_vpfs

    eq_cols_all = [f"{c}__equal_hsbc" for c in SPECTRA_TO_HSBC_MAP.keys()] + [f"{c}__equal_vpfs" for c in SPECTRA_TO_HSBC_MAP.keys()]
    df["has_diff"] = ~df[eq_cols_all].all(axis=1)
    return df


def build_triple_adjacent_export_df(comp: pd.DataFrame) -> pd.DataFrame:
    work = comp.copy()
    # 统一前置列
    if "spectra security ID" not in work.columns:
        work = work.assign(**{
            "spectra security ID": (work["id_value"] if "id_value" in work.columns else pd.Series(dtype=object))
        })
    if "HSBC security ID" not in work.columns and "Security ID" in work.columns:
        work = work.assign(**{
            "HSBC security ID": work["Security ID"]
        })

    front_cols = [c for c in [
        "id_type", "id_value", "_type_raw",
        "spectra security ID", "HSBC security ID", "VPFS_security ID", "Isin", "Ticker"
    ] if c in work.columns]
    ordered_cols: list[str] = front_cols.copy()

    for spectra_col in SPECTRA_TO_HSBC_MAP.keys():
        left_alias = f"{spectra_col}__spectra"
        hsbc_alias = f"{spectra_col}__hsbc"
        vpfs_alias = f"{spectra_col}__vpfs"
        hsbc_col = SPECTRA_TO_HSBC_MAP[spectra_col]
        vpfs_col = SPECTRA_TO_VPFS_MAP[spectra_col]
        # 确保基础列存在
        for col in [spectra_col, hsbc_col, vpfs_col]:
            if col not in work.columns:
                work.loc[:, col] = pd.Series(dtype=object)
        # 生成别名列
        work[left_alias] = work[spectra_col]
        work[hsbc_alias] = work[hsbc_col]
        work[vpfs_alias] = work[vpfs_col]
        # 追加顺序：左、HSBC、VPFS、两个 delta（若有）、两个 equal（若有）
        ordered_cols += [left_alias, hsbc_alias, vpfs_alias]
        for suffix in ["delta_hsbc", "delta_vpfs", "equal_hsbc", "equal_vpfs"]:
            coln = f"{spectra_col}__{suffix}"
            if coln in work.columns:
                ordered_cols.append(coln)

    if "has_diff" in work.columns:
        ordered_cols.append("has_diff")

    seen = set()
    final_cols = []
    for c in ordered_cols:
        if c in work.columns and c not in seen:
            seen.add(c)
            final_cols.append(c)
    return work[final_cols].copy()


def build_triple_diff_columns_series(comp: pd.DataFrame) -> pd.Series:
    df = comp.copy()
    names = list(SPECTRA_TO_HSBC_MAP.keys())
    for col in [f"{n}__equal_hsbc" for n in names] + [f"{n}__equal_vpfs" for n in names]:
        if col not in df.columns:
            df.loc[:, col] = False
    def row_to_list(r):
        diffs = []
        for n in names:
            if not bool(r.get(f"{n}__equal_hsbc", False)):
                diffs.append(f"{n}(hsbc)")
            if not bool(r.get(f"{n}__equal_vpfs", False)):
                diffs.append(f"{n}(vpfs)")
        return ", ".join(diffs)
    return df.apply(row_to_list, axis=1)


def build_triple_all_sheets_bytes(export_comp: pd.DataFrame, export_diffs: pd.DataFrame, export_unmatched: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        export_comp.to_excel(writer, sheet_name="comparison", index=False)
        export_unmatched.to_excel(writer, sheet_name="unmatched", index=False)
        sheet = "diffs"
        export_diffs.to_excel(writer, sheet_name=sheet, index=False)
        if not export_diffs.empty:
            workbook = writer.book
            worksheet = writer.sheets[sheet]
            yellow = workbook.add_format({"bg_color": "#FFFF00"})
            rows = len(export_diffs)
            for spectra_col in SPECTRA_TO_HSBC_MAP.keys():
                left_alias = f"{spectra_col}__spectra"
                for side in ["hsbc", "vpfs"]:
                    right_alias = f"{spectra_col}__{side}"
                    equal_col = f"{spectra_col}__equal_{side}"
                    if left_alias not in export_diffs.columns or right_alias not in export_diffs.columns or equal_col not in export_diffs.columns:
                        continue
                    left_idx = export_diffs.columns.get_loc(left_alias)
                    right_idx = export_diffs.columns.get_loc(right_alias)
                    equal_idx = export_diffs.columns.get_loc(equal_col)
                    equal_letter = _col_idx_to_excel(equal_idx)
                    for r in range(rows):
                        excel_row = r + 2
                        formula = f"=${equal_letter}{excel_row}=FALSE"
                        row_idx = r + 1
                        worksheet.conditional_format(row_idx, left_idx, row_idx, left_idx, {
                            "type": "formula",
                            "criteria": formula,
                            "format": yellow,
                        })
                        worksheet.conditional_format(row_idx, right_idx, row_idx, right_idx, {
                            "type": "formula",
                            "criteria": formula,
                            "format": yellow,
                        })
    buf.seek(0)
    return buf.getvalue()


def run_compare_triple_from_sources(spectra_src: Path | BytesIO, hsbc_src: Path | BytesIO, vpfs_src: Path | BytesIO) -> dict:
    # 左：Spectra
    spectra_df = read_spectra_raw_from(spectra_src)
    spectra_norm = spectra_normalize(spectra_df)
    # 右1：HSBC（含 fallback）
    hsbc_df = read_hsbc_raw_from(hsbc_src)
    long_df, dups, hsbc_base = hsbc_build_long_table(hsbc_df)
    merged, _ = left_join_non_dup(spectra_norm, long_df, dups)
    merged_fb, _, _ = apply_fallback_mapping(merged, hsbc_base)
    # 右2：VPFS
    vpfs_df = read_vpfs_raw_from(vpfs_src)
    vpfs_norm = vpfs_normalize(vpfs_df)
    vpfs_triple = _vpfs_for_triple(vpfs_norm)
    # 合并 VPFS（避免覆盖 HSBC 列）
    merged_triple = merged_fb.merge(
        vpfs_triple,
        left_on="id_value",
        right_on="prodid_norm",
        how="left",
        suffixes=(None, None),
    )

    triple_comp = build_triple_comparison(merged_triple)
    export_comp = build_triple_adjacent_export_df(triple_comp)
    export_comp.loc[:, "diff_columns"] = build_triple_diff_columns_series(triple_comp)
    export_diffs = export_comp[export_comp.get("has_diff", pd.Series([False] * len(export_comp))).astype(bool)].copy()

    # 未匹配（左）：HSBC 四列全空 或 VPFS 四列全空
    hsbc_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in merged_triple.columns]
    vpfs_cols = [v for v in SPECTRA_TO_VPFS_MAP.values() if v in merged_triple.columns]
    left_unmatched_frames = []
    if hsbc_cols:
        mask_hsbc_unmatched = merged_triple[hsbc_cols].isna().all(axis=1)
        left_unmatched_frames.append(merged_triple[mask_hsbc_unmatched].assign(source="left_vs_hsbc"))
    if vpfs_cols:
        mask_vpfs_unmatched = merged_triple[vpfs_cols].isna().all(axis=1)
        left_unmatched_frames.append(merged_triple[mask_vpfs_unmatched].assign(source="left_vs_vpfs"))
    export_unmatched_left = pd.concat(left_unmatched_frames, ignore_index=True, sort=False) if left_unmatched_frames else merged_triple.iloc[0:0].copy()
    export_unmatched_left = build_triple_adjacent_export_df(build_triple_comparison(export_unmatched_left.copy()) if not export_unmatched_left.empty else export_unmatched_left.copy()) if not export_unmatched_left.empty else export_unmatched_left
    if not export_unmatched_left.empty:
        export_unmatched_left.loc[:, "diff_columns"] = build_triple_diff_columns_series(build_triple_comparison(export_unmatched_left.copy()))

    # 未覆盖（右）：HSBC base 的 Security ID 未出现在 comparison，及 VPFS 的 prodid_norm 未出现在 spectra id_value
    # HSBC 右侧未覆盖
    comp_hsbc_ids = set(export_comp.get("HSBC security ID", pd.Series(dtype=str)).astype(str).str.strip().str.upper().dropna().tolist())
    if "Security ID" in hsbc_base.columns:
        hsbc_ids_all = hsbc_base["Security ID"].astype(str).str.strip().str.upper()
        rhs_missing_hsbc = hsbc_base[~hsbc_ids_all.isin(comp_hsbc_ids)].copy()
    else:
        rhs_missing_hsbc = hsbc_base.iloc[0:0].copy()
    # VPFS 右侧未覆盖
    spectra_ids = set(export_comp.get("spectra security ID", pd.Series(dtype=str)).astype(str).str.strip().str.upper().dropna().tolist())
    if "prodid_norm" in vpfs_triple.columns:
        vpfs_only = vpfs_triple[~vpfs_triple["prodid_norm"].astype(str).str.strip().str.upper().isin(spectra_ids)].copy()
    else:
        vpfs_only = vpfs_triple.iloc[0:0].copy()

    comp_rhs_hsbc_only = build_triple_comparison(rhs_missing_hsbc.copy()) if not rhs_missing_hsbc.empty else rhs_missing_hsbc.copy()
    export_unmatched_right_hsbc = build_triple_adjacent_export_df(comp_rhs_hsbc_only) if not comp_rhs_hsbc_only.empty else comp_rhs_hsbc_only.copy()
    if not export_unmatched_right_hsbc.empty:
        export_unmatched_right_hsbc.loc[:, "diff_columns"] = build_triple_diff_columns_series(comp_rhs_hsbc_only)
        export_unmatched_right_hsbc.loc[:, "source"] = "right_hsbc"

    comp_rhs_vpfs_only = build_triple_comparison(vpfs_only.copy()) if not vpfs_only.empty else vpfs_only.copy()
    export_unmatched_right_vpfs = build_triple_adjacent_export_df(comp_rhs_vpfs_only) if not comp_rhs_vpfs_only.empty else comp_rhs_vpfs_only.copy()
    if not export_unmatched_right_vpfs.empty:
        export_unmatched_right_vpfs.loc[:, "diff_columns"] = build_triple_diff_columns_series(comp_rhs_vpfs_only)
        export_unmatched_right_vpfs.loc[:, "source"] = "right_vpfs"

    export_unmatched = pd.concat([f for f in [export_unmatched_left, export_unmatched_right_hsbc, export_unmatched_right_vpfs] if isinstance(f, pd.DataFrame) and not f.empty], ignore_index=True, sort=False) if (not export_unmatched_left.empty or (isinstance(export_unmatched_right_hsbc, pd.DataFrame) and not export_unmatched_right_hsbc.empty) or (isinstance(export_unmatched_right_vpfs, pd.DataFrame) and not export_unmatched_right_vpfs.empty)) else pd.DataFrame(columns=export_comp.columns)

    all_bytes = build_triple_all_sheets_bytes(export_comp, export_diffs, export_unmatched)
    return {
        "comparison": export_comp,
        "diffs": export_diffs,
        "unmatched": export_unmatched,
        "duplicates": pd.DataFrame(),  # 三源模式暂不输出重复键
        "all_sheets_xlsx": all_bytes,
    }
if __name__ == "__main__":
    try:
        # 顺序执行步骤，边实现边测试
        rc = main_step1_read_files()
        if rc != 0:
            sys.exit(rc)

        print("[Step 2] 构建 HSBC 长表并检测重复键…")
        hsbc_df = read_hsbc_raw()
        long_df, dups, hsbc_base = hsbc_build_long_table(hsbc_df)
        print(f"hsbc long: shape={long_df.shape}")
        print("hsbc long columns:", list(long_df.columns))
        print(f"hsbc duplicate keys: {dups[['id_type','id_value']].drop_duplicates().shape[0]}")
        if not dups.empty:
            print("示例重复键前5行:")
            print(dups[["id_type", "id_value", "Security ID"]].head(5).to_string(index=False))

        print("[Step 2] OK")
        print("[Step 3] 左连接（排除 dup 键）并打印计数…")
        spectra_df = read_spectra_raw()
        spectra_norm = spectra_normalize(spectra_df)
        merged, unmatched = left_join_non_dup(spectra_norm, long_df, dups)
        print(f"merged: shape={merged.shape}, unmatched: {unmatched.shape[0]}")
        print("合并列(前30):", list(merged.columns[:30]))
        print("合并示例前5行:")
        sample_cols = [
            c for c in [
                "id_type", "id_value", "_type_raw",
                "Shares/Par", "Price", "Traded Market Value", "Traded Market Value (Base)",
                "Security ID", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value",
            ] if c in merged.columns
        ]
        print(merged[sample_cols].head(5).to_string(index=False))
        print("[Step 3] OK")

        print("[Step 4] 严格比较四列并统计差异…")
        comp = build_comparison(merged)
        diff_count = int(comp["has_diff"].sum())
        print(f"diff rows: {diff_count}")
        # 打印前 5 条差异样例
        diff_cols_print = [
            "id_type", "id_value", "_type_raw",
            "Shares/Par", "Quantity", "Shares/Par__delta", "Shares/Par__equal",
            "Price", "Local Market Price", "Price__delta", "Price__equal",
            "Traded Market Value", "Local Market Value", "Traded Market Value__delta", "Traded Market Value__equal",
            "Traded Market Value (Base)", "Book Market Value", "Traded Market Value (Base)__delta", "Traded Market Value (Base)__equal",
        ]
        diff_cols_print = [c for c in diff_cols_print if c in comp.columns]
        diffs_preview = comp[comp["has_diff"]][diff_cols_print].head(5)
        if not diffs_preview.empty:
            print("差异示例前5行:")
            print(diffs_preview.to_string(index=False))
        print("[Step 4] OK")

        print("[Step 4b] 对未匹配行应用 fallback 映射并重算…")
        merged_fallback, updated_ct, dup_skipped = apply_fallback_mapping(merged, hsbc_base)
        print(f"fallback 更新匹配数: {updated_ct}, 因 Security ID 多行被跳过: {dup_skipped}")
        comp = build_comparison(merged_fallback)
        diff_count = int(comp["has_diff"].sum())
        # 重新计算未匹配（基于 hsbc 四列是否全空）
        hsbc_cols_check = [c for c in ["Security ID", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in merged_fallback.columns]
        if hsbc_cols_check:
            unmatched_mask_fb = merged_fallback[hsbc_cols_check].isna().all(axis=1)
        else:
            unmatched_mask_fb = pd.Series([True] * len(merged_fallback), index=merged_fallback.index)
        unmatched_fb = merged_fallback[unmatched_mask_fb]
        print(f"fallback 后未匹配: {unmatched_fb.shape[0]}, 差异行: {diff_count}")
        print("[Step 4b] OK")

        print("[Step 5] 导出 CSV…")
        # 产出文件：comparison.csv、diffs.csv、unmatched.csv、duplicates.csv、summary.csv
        out_comparison = Path("comparison.csv")
        out_diffs = Path("diffs.csv")
        out_unmatched = Path("unmatched.csv")
        out_duplicates = Path("duplicates.csv")
        out_summary = Path("summary.csv")

        export_comp = build_adjacent_export_df(comp)
        export_diffs = export_comp[export_comp["has_diff"]].copy() if "has_diff" in export_comp.columns else export_comp.iloc[0:0].copy()
        # 追加 diff_columns
        export_comp.loc[:, "diff_columns"] = build_diff_columns_series(comp)
        if not export_diffs.empty:
            mask_has_diff = export_comp["has_diff"] if "has_diff" in export_comp.columns else pd.Series([False] * len(export_comp))
            export_diffs.loc[:, "diff_columns"] = build_diff_columns_series(comp[mask_has_diff].copy())
        export_comp.to_csv(out_comparison, index=False)
        export_diffs.to_csv(out_diffs, index=False)
        # 额外导出带黄色高亮差异的 Excel
        out_diffs_xlsx = Path("diffs_highlight.xlsx")
        write_diffs_excel(export_diffs, out_diffs_xlsx)
        # unmatched：从 comp 中筛选 hsbc 侧四列全空的行，再做相邻导出
        hsbc_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in merged_fallback.columns]
        if hsbc_cols:
            unmatched_mask_export = merged_fallback[hsbc_cols].isna().all(axis=1)
        else:
            unmatched_mask_export = pd.Series([True] * len(merged_fallback), index=merged_fallback.index)
        left_unmatched = merged_fallback[unmatched_mask_export].copy()
        export_unmatched_left = build_adjacent_export_df(left_unmatched)
        export_unmatched_left.loc[:, "diff_columns"] = build_diff_columns_series(build_comparison(left_unmatched.copy()))
        export_unmatched_left.loc[:, "source"] = "left"

        comp_hsbc_ids = set(export_comp.get("HSBC security ID", pd.Series(dtype=str)).astype(str).str.strip().str.upper().dropna().tolist())
        if "Security ID" in hsbc_base.columns:
            hsbc_ids_all = hsbc_base["Security ID"].astype(str).str.strip().str.upper()
            rhs_missing = hsbc_base[~hsbc_ids_all.isin(comp_hsbc_ids)].copy()
        else:
            rhs_missing = hsbc_base.iloc[0:0].copy()
        comp_rhs_only = build_comparison(rhs_missing.copy()) if not rhs_missing.empty else rhs_missing.copy()
        export_unmatched_right = build_adjacent_export_df(comp_rhs_only) if not comp_rhs_only.empty else comp_rhs_only.copy()
        if not export_unmatched_right.empty:
            export_unmatched_right.loc[:, "diff_columns"] = build_diff_columns_series(comp_rhs_only)
            export_unmatched_right.loc[:, "source"] = "right"

        export_unmatched = pd.concat([export_unmatched_left, export_unmatched_right], ignore_index=True, sort=False)
        export_unmatched.to_csv(out_unmatched, index=False)
        # duplicates：将 dups 视为仅有 HSBC 值的比较输入，补齐 spectra 列为 None 后构造相邻导出
        dups_for_comp = dups.copy()
        dups_for_comp = dups_for_comp[[c for c in ["id_type", "id_value", "Security ID", "Isin", "Ticker", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in dups_for_comp.columns]]
        dups_comp = build_comparison(dups_for_comp)
        export_dups = build_adjacent_export_df(dups_comp)
        export_dups.loc[:, "diff_columns"] = build_diff_columns_series(dups_comp)
        export_dups.to_csv(out_duplicates, index=False)

        # summary：计数与关键清单引用（以简单的 CSV 形式输出几行统计）
        total_rows = comp.shape[0]
        matched_rows = total_rows - unmatched.shape[0]
        diff_rows = int(comp["has_diff"].sum())
        dup_keys = dups[["id_type", "id_value"]].drop_duplicates().shape[0]
        summary_rows = [
            {"metric": "total", "value": total_rows},
            {"metric": "matched", "value": matched_rows},
            {"metric": "unmatched", "value": unmatched.shape[0]},
            {"metric": "diff_rows", "value": diff_rows},
            {"metric": "duplicate_keys", "value": dup_keys},
            {"metric": "files", "value": "comparison.csv, diffs.csv, unmatched.csv, duplicates.csv"},
        ]
        pd.DataFrame(summary_rows).to_csv(out_summary, index=False)
        print(f"导出完成: {out_comparison}, {out_diffs}, {out_unmatched}, {out_duplicates}, {out_summary}")
        print("[Step 5] OK")
        sys.exit(0)
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)


