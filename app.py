import os
from typing import Optional

import pandas as pd
import streamlit as st

from compare import (
    LEFT_SHEET_DEFAULT,
    RIGHT_SHEET_DEFAULT,
    DEFAULT_COLUMN_MAPPING,
    suggest_keys,
    read_excel_sheet,
    CompareConfig,
    compare_data,
    export_to_excel_bytes,
    compare_data_type_rules,
)


st.set_page_config(page_title="Excel 双表比对工具", layout="wide")
st.title("Excel 双表比对工具 - 临时可视化版")
st.caption("拖拽上传两个 Excel，或填写路径。默认使用指定 sheet，可选类型规则或自动主键，严格等值比对四个数值列。")


with st.sidebar:
    st.header("参数设置")
    left_sheet = st.text_input("左侧 sheet 名称", value=LEFT_SHEET_DEFAULT)
    right_sheet = st.text_input("右侧 sheet 名称", value=RIGHT_SHEET_DEFAULT)
    join_mode = st.radio("主键模式", options=["类型规则匹配", "自动主键"], index=0)
    st.markdown("**列映射（固定）**")
    for l, r in DEFAULT_COLUMN_MAPPING.items():
        st.write(f"{l} ↔ {r}")


col1, col2 = st.columns(2)
with col1:
    left_file = st.file_uploader(
        "上传左侧 Excel（Security Distribution）", type=["xlsx", "xls"], key="left_upl"
    )
    left_path = st.text_input("或输入左侧文件路径（可留空）")
with col2:
    right_file = st.file_uploader(
        "上传右侧 Excel（HSBC Position Appraisal Report）", type=["xlsx", "xls"], key="right_upl"
    )
    right_path = st.text_input("或输入右侧文件路径（可留空）")


def _resolve_source(file_uploader, path_input: str) -> Optional[object]:
    if file_uploader is not None:
        return file_uploader
    if path_input:
        if os.path.exists(path_input):
            return path_input
        else:
            st.warning(f"路径不存在：{path_input}")
            return None
    return None


left_src = _resolve_source(left_file, left_path)
right_src = _resolve_source(right_file, right_path)

if (left_src is not None) and (right_src is not None):
    try:
        left_df = read_excel_sheet(left_src, sheet_name=left_sheet)
        right_df = read_excel_sheet(right_src, sheet_name=right_sheet)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success("文件读取成功。")

    # 展示预览
    with st.expander("查看前20行预览（左/右）"):
        st.write("左侧：")
        st.dataframe(left_df.head(20))
        st.write("右侧：")
        st.dataframe(right_df.head(20))

    left_key = None
    right_key = None

    if join_mode == "自动主键":
        candidates = suggest_keys(left_df, right_df)
        if not candidates:
            st.warning("未找到合适的候选主键。请在下方手工输入列名。")
            left_key = st.text_input("左侧主键列名")
            right_key = st.text_input("右侧主键列名")
        else:
            st.subheader("候选主键（按推荐度排序）")
            cand_df = pd.DataFrame(
                [
                    {
                        "left_key": c.left_key,
                        "right_key": c.right_key,
                        "coverage_ratio": round(c.coverage_ratio, 4),
                        "left_unique_ratio": round(c.left_unique_ratio, 4),
                        "right_unique_ratio": round(c.right_unique_ratio, 4),
                        "matched_rows": c.matched_rows,
                        "score": round(c.score, 4),
                    }
                    for c in candidates
                ]
            )
            st.dataframe(cand_df)

            best = candidates[0]
            left_key = st.selectbox(
                "选择左侧主键列",
                options=[c.left_key for c in candidates],
                index=[c.left_key for c in candidates].index(best.left_key),
            )
            options_for_left = [c.right_key for c in candidates if c.left_key == left_key]
            right_key = st.selectbox(
                "选择右侧主键列",
                options=options_for_left or [c.right_key for c in candidates],
                index=0,
            )

    run = st.button("开始比对", type="primary")
    if run:
        try:
            if join_mode == "类型规则匹配":
                result = compare_data_type_rules(left_df, right_df, DEFAULT_COLUMN_MAPPING)
            else:
                if not left_key or not right_key:
                    st.error("请先选择或填写主键列名。")
                    st.stop()
                config = CompareConfig(
                    left_key=left_key,
                    right_key=right_key,
                    column_mapping=DEFAULT_COLUMN_MAPPING,
                )
                result = compare_data(left_df, right_df, config)
        except Exception as e:
            st.error(f"比对出错：{e}")
            st.stop()

        st.subheader("摘要")
        st.dataframe(result.summary)
        if getattr(result, "notes", None):
            st.caption(result.notes)

        st.subheader("差异明细（仅显示有差异的匹配行）")
        st.dataframe(result.diffs)

        with st.expander("仅在左侧/右侧存在的主键"):
            if join_mode == "类型规则匹配":
                st.dataframe(result.only_in_left.head(200))
                st.dataframe(result.only_in_right.head(200))
            else:
                left_cols = [c for c in [left_key] + list(DEFAULT_COLUMN_MAPPING.keys()) if c in result.only_in_left.columns]
                right_cols = [c for c in [right_key] + list(DEFAULT_COLUMN_MAPPING.values()) if c in result.only_in_right.columns]
                st.write("仅在左侧：")
                st.dataframe(result.only_in_left[left_cols].head(200))
                st.write("仅在右侧：")
                st.dataframe(result.only_in_right[right_cols].head(200))

        excel_bytes = export_to_excel_bytes(result)
        st.download_button(
            label="下载对比结果 Excel",
            data=excel_bytes,
            file_name="compare_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("请上传或填写两个 Excel 文件后开始。")
