import streamlit as st
import pandas as pd
from io import BytesIO

from compare import run_compare_from_sources, build_single_sheet_excel, SPECTRA_TO_HSBC_MAP


st.set_page_config(page_title="Excel Compare", layout="wide")
st.title("Excel Compare Tool")

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

    tabs = st.tabs(["comparison", "diffs", "unmatched", "duplicates"]) 
    for tab_name, tab in zip(["comparison", "diffs", "unmatched", "duplicates"], tabs):
        with tab:
            df: pd.DataFrame = result[tab_name]
            if tab_name == "diffs" and not df.empty:
                # Web 端高亮：使用 pandas Styler 生成 HTML，按 equal 列对左右值列染色
                def style_diffs(d: pd.DataFrame):
                    styler = d.style
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
                        if equal_col in d.columns:
                            cond = ~d[equal_col].fillna(False)
                            colors = ["background-color: #ffff00" if v else "" for v in cond.tolist()]
                            if left_alias in d.columns:
                                styler = styler.apply(lambda s, c=colors: c, subset=[left_alias])
                            if right_alias in d.columns:
                                styler = styler.apply(lambda s, c=colors: c, subset=[right_alias])
                    try:
                        styler = styler.format(precision=6)
                    except Exception:
                        pass
                    return styler

                styled = style_diffs(df)
                st.markdown(styled.to_html(), unsafe_allow_html=True)
            else:
                st.dataframe(df, use_container_width=True)

            xlsx_bytes = build_single_sheet_excel(df, sheet_name=tab_name)
            st.download_button(
                label=f"Download {tab_name} (Excel)",
                data=xlsx_bytes,
                file_name=f"{tab_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.divider()
    st.download_button(
        label="Download all sheets (Excel with Diff Highlight)",
        data=result["all_sheets_xlsx"],
        file_name="comparison_all.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


