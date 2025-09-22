import streamlit as st
import pandas as pd
from io import BytesIO

from compare import run_compare_from_sources, build_single_sheet_excel, SPECTRA_TO_HSBC_MAP


st.set_page_config(page_title="Excel Compare", layout="wide")
st.title("Excel Compare Tool")

# 持久化页面状态，避免任意按钮导致回到上传视图
if "view" not in st.session_state:
    st.session_state["view"] = "compare"
if "result" not in st.session_state:
    st.session_state["result"] = None

if st.session_state["view"] == "compare":
    col1, col2 = st.columns(2)
    with col1:
        spectra_file = st.file_uploader("Upload spectra.xls", type=["xls", "xlsx"]) 
    with col2:
        hsbc_file = st.file_uploader("Upload HSBC Position Appraisal Report (EXCEL).xlsx", type=["xlsx", "xls"]) 

    run_btn = st.button("Run Compare", disabled=not (spectra_file and hsbc_file))

    if run_btn and spectra_file and hsbc_file:
        spectra_bytes = BytesIO(spectra_file.read())
        hsbc_bytes = BytesIO(hsbc_file.read())
        with st.spinner("Comparing…"):
            result = run_compare_from_sources(spectra_bytes, hsbc_bytes)
        st.session_state["result"] = result
        st.session_state["view"] = "results"
        st.rerun()
else:
    # 左上角返回按钮：仅当用户主动点击时回到上传页；其他按钮不会改变视图
    back_clicked = st.button("← 返回", key="back_to_compare")
    if back_clicked:
        st.session_state["view"] = "compare"
        st.session_state["result"] = None
        st.rerun()

    result = st.session_state["result"] or {}

    tabs = st.tabs(["comparison", "diffs", "unmatched", "duplicates"]) 
    for tab_name, tab in zip(["comparison", "diffs", "unmatched", "duplicates"], tabs):
        with tab:
            df: pd.DataFrame = result.get(tab_name, pd.DataFrame())
            if tab_name == "diffs" and not df.empty:
                # 紧凑展示（固定参数）：显示所有列、紧凑密度、冻结首列、固定高度
                compact_mode = True
                table_height = 520
                view_df = df

                # Web 端高亮：按 base_df 的 equal 列对左右值列染色
                def style_diffs(view: pd.DataFrame, base: pd.DataFrame):
                    styler = view.style
                    # 隐藏索引（兼容不同 pandas 版本）
                    try:
                        styler = styler.hide(axis="index")
                    except Exception:
                        try:
                            styler = styler.hide_index()
                        except Exception:
                            pass
                    for field in SPECTRA_TO_HSBC_MAP.keys():
                        left_alias = f"{field}__spectra"
                        right_alias = f"{field}__hsbc"
                        equal_col = f"{field}__equal"
                        if equal_col in base.columns:
                            cond = ~base[equal_col].fillna(False)
                            colors = ["background-color: #ffff00" if v else "" for v in cond.tolist()]
                            if left_alias in view.columns:
                                styler = styler.apply(lambda s, c=colors: c, subset=[left_alias])
                            if right_alias in view.columns:
                                styler = styler.apply(lambda s, c=colors: c, subset=[right_alias])
                    try:
                        styler = styler.format(precision=6)
                    except Exception:
                        pass
                    return styler

                styled = style_diffs(view_df, df)

                css = f"""
                <style>
                .diff-container {{
                    max-height: {table_height}px;
                    overflow: auto;
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                }}
                .diff-container table {{
                    width: 100%;
                    border-collapse: collapse;
                }}
                .diff-container thead th {{
                    position: sticky;
                    top: 0;
                    background: #ffffff;
                    z-index: 2;
                    box-shadow: 0 1px 0 rgba(0,0,0,0.06);
                }}
                .diff-container td, .diff-container th {{
                    padding: {"4px 8px" if compact_mode else "8px 10px"};
                    font-size: {"12px" if compact_mode else "14px"};
                    white-space: nowrap;
                }}
                .diff-container tbody tr:nth-child(even) td {{
                    background-color: #fafafa;
                }}
                .diff-container tbody td:first-child,
                .diff-container thead th:first-child {{
                    position: sticky;
                    left: 0;
                    background: #ffffff;
                    z-index: 3;
                    box-shadow: 1px 0 0 rgba(0,0,0,0.04);
                }}
                </style>
                """
                st.markdown(css, unsafe_allow_html=True)
                st.markdown(f'<div class="diff-container">{styled.to_html()}</div>', unsafe_allow_html=True)
            else:
                st.dataframe(df, use_container_width=True)

            if not df.empty:
                xlsx_bytes = build_single_sheet_excel(df, sheet_name=tab_name)
                st.download_button(
                    label=f"Download {tab_name} (Excel)",
                    data=xlsx_bytes,
                    file_name=f"{tab_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.divider()
    if "all_sheets_xlsx" in result:
        st.download_button(
            label="Download all sheets (Excel with Diff Highlight)",
            data=result["all_sheets_xlsx"],
            file_name="comparison_all.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


