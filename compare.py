import sys
from pathlib import Path

import pandas as pd
from decimal import Decimal, InvalidOperation

from config import (
    bond_code_set,
    stack_code_set,
    missing_isin_or_stack_code_mapping_dict,
)


SPECTRA_PATH = Path("./spectra.xls")
SPECTRA_SHEET = "Security Distribution"

HSBC_PATH = Path("./HSBC Position Appraisal Report (EXCEL).xlsx")
HSBC_SHEET = "HSBC Position Appraisal Report"


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


def hsbc_build_long_table(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
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
    return long_df, dups


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


def build_comparison(merged: pd.DataFrame) -> pd.DataFrame:
    df = merged.copy()
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
        eq = lhs_num == rhs_num
        df[f"{spectra_col}__equal"] = eq
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
    front_cols = [c for c in ["id_type", "id_value", "_type_raw", "Security ID", "Isin", "Ticker"] if c in comp.columns]
    ordered_cols: list[str] = front_cols.copy()

    for spectra_col, hsbc_col in SPECTRA_TO_HSBC_MAP.items():
        # 创建别名列，确保存在
        left_alias = f"{spectra_col}__spectra"
        right_alias = f"{spectra_col}__hsbc"
        delta_col = f"{spectra_col}__delta"
        equal_col = f"{spectra_col}__equal"

        if spectra_col not in comp.columns:
            comp[spectra_col] = None
        if hsbc_col not in comp.columns:
            comp[hsbc_col] = None

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


if __name__ == "__main__":
    try:
        # 顺序执行步骤，边实现边测试
        rc = main_step1_read_files()
        if rc != 0:
            sys.exit(rc)

        print("[Step 2] 构建 HSBC 长表并检测重复键…")
        hsbc_df = read_hsbc_raw()
        long_df, dups = hsbc_build_long_table(hsbc_df)
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

        print("[Step 5] 导出 CSV…")
        # 产出文件：comparison.csv、diffs.csv、unmatched.csv、duplicates.csv、summary.csv
        out_comparison = Path("comparison.csv")
        out_diffs = Path("diffs.csv")
        out_unmatched = Path("unmatched.csv")
        out_duplicates = Path("duplicates.csv")
        out_summary = Path("summary.csv")

        export_comp = build_adjacent_export_df(comp)
        export_diffs = export_comp[export_comp["has_diff"]] if "has_diff" in export_comp.columns else export_comp.iloc[0:0]
        export_comp.to_csv(out_comparison, index=False)
        export_diffs.to_csv(out_diffs, index=False)
        # unmatched：从 comp 中筛选 hsbc 侧四列全空的行，再做相邻导出
        hsbc_cols = [v for v in SPECTRA_TO_HSBC_MAP.values() if v in comp.columns]
        if hsbc_cols:
            unmatched_mask_export = comp[hsbc_cols].isna().all(axis=1)
        else:
            unmatched_mask_export = pd.Series([True] * len(comp), index=comp.index)
        export_unmatched = build_adjacent_export_df(comp[unmatched_mask_export].copy())
        export_unmatched.to_csv(out_unmatched, index=False)
        # duplicates：将 dups 视为仅有 HSBC 值的比较输入，补齐 spectra 列为 None 后构造相邻导出
        dups_for_comp = dups.copy()
        dups_for_comp = dups_for_comp[[c for c in ["id_type", "id_value", "Security ID", "Isin", "Ticker", "Quantity", "Local Market Price", "Local Market Value", "Book Market Value"] if c in dups_for_comp.columns]]
        dups_comp = build_comparison(dups_for_comp)
        export_dups = build_adjacent_export_df(dups_comp)
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


