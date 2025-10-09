import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from pathlib import Path

import config
from compare import run_compare_from_sources, build_single_sheet_excel, SPECTRA_TO_HSBC_MAP, run_compare_spectra_vs_vpfs, run_compare_triple_from_sources


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

    # VPFS 上传入口（d01 2.xls）
    st.divider()
    vpfs_file = st.file_uploader("Upload VPFS (d01 2.xls)", type=["xls", "xlsx"], key="vpfs_uploader")

    # 显示并编辑当前的 missing_isin_or_stack_code_mapping_dict
    st.subheader("Security 映射管理")
    with st.expander("missing_isin_or_stack_code_mapping_dict（增删改查）", expanded=True):
        # 上次操作后的网页提示（toast），通过 session_state 传递跨 rerun 显示
        if "notify" in st.session_state:
            _n = st.session_state.pop("notify")
            _msg = _n.get("message", "")
            _typ = _n.get("type", "info")
            _icon = "✅" if _typ == "add" else ("🗑️" if _typ == "delete" else "ℹ️")
            st.toast(_msg, icon=_icon)
        st.markdown(
            "- 在下方‘搜索’中输入关键字进行【查】（按 Key/Value 模糊匹配）\n"
            "- 直接在表格中修改 Value 完成【改】；可新增行完成【增】\n"
            "- 勾选‘删除’列并点击‘删除勾选项’完成【删】\n"
            "- 点击‘保存’写入 mapping_override.json 并立即生效"
        )
        current_map = config.missing_isin_or_stack_code_mapping_dict
        # 搜索过滤（大小写不敏感）
        search_text = st.text_input("搜索 Key/Value（模糊匹配，忽略大小写）", key="mapping_search")
        # 构造完整表（将 delete 放到最后，便于稳定缩窄宽度）
        full_items = [{"key": k, "value": v, "delete": False} for k, v in current_map.items()]
        full_df = pd.DataFrame(full_items, columns=["key", "value", "delete"]) if full_items else pd.DataFrame(columns=["key", "value", "delete"]) 
        # 过滤视图
        if search_text:
            s = str(search_text).strip().upper()
            mask = full_df.apply(lambda r: (str(r.get("key", "")).upper().find(s) >= 0) or (str(r.get("value", "")).upper().find(s) >= 0), axis=1)
            view_df_src = full_df[mask].copy()
        else:
            view_df_src = full_df.copy()
        st.caption(f"共 {len(full_df)} 条；当前显示 {len(view_df_src)} 条")
        # 收窄 delete 列宽（尽量贴近复选框宽度）
        st.markdown(
            """
            <style>
            [data-testid=\"stDataEditor\"] table thead th:last-child,
            [data-testid=\"stDataEditor\"] table tbody td:last-child {
                width: 52px;
                max-width: 52px;
                min-width: 44px;
                padding-left: 8px;
                padding-right: 8px;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        # 渲染可编辑
        edited_df = st.data_editor(
            view_df_src,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="mapping_editor",
            column_config={
                "key": st.column_config.TextColumn(
                    "Key",
                ),
                "value": st.column_config.TextColumn(
                    "Value",
                ),
                "delete": st.column_config.CheckboxColumn(
                    "删",
                    help="勾选后点击‘删除勾选项’",
                    default=False,
                ),
            },
        )
        # 新增表单（单独入口，强提示）
        col_add1, col_add2, col_add3 = st.columns([2,2,1])
        with col_add1:
            new_key = st.text_input("新增 Key（不区分大小写）", key="new_map_key")
        with col_add2:
            new_val = st.text_input("新增 Value（HSBC Security ID）", key="new_map_val")
        with col_add3:
            add_entry = st.button("添加", key="add_mapping_btn")

        col_ops1, col_ops2 = st.columns([1,1])
        with col_ops1:
            del_selected = st.button("删除勾选项", key="delete_selected_btn")
        with col_ops2:
            save_mapping = st.button("保存", key="save_mapping_btn")

        # 将视图编辑结果合并回完整表（避免过滤导致未显示行丢失）
        def merge_view_to_full(full_df_in: pd.DataFrame, view_df_edited: pd.DataFrame) -> pd.DataFrame:
            if not isinstance(view_df_edited, pd.DataFrame):
                return full_df_in.copy()
            # 统一键格式
            view_df_edited = view_df_edited.copy()
            view_df_edited["key"] = view_df_edited["key"].astype(str).str.strip()
            # 剔除空 key
            view_df_edited = view_df_edited[view_df_edited["key"] != ""]
            # 视图中的键集合
            view_keys = set(view_df_edited["key"].str.upper().tolist())
            # 保留视图外的原始行
            rest = full_df_in[~full_df_in["key"].astype(str).str.upper().isin(view_keys)].copy()
            # 合并：视图编辑行 + 其他行
            merged = pd.concat([rest, view_df_edited], ignore_index=True)
            return merged

        # 处理『增』
        if add_entry:
            try:
                k = (new_key or "").strip()
                v = (new_val or "").strip()
                if not k:
                    st.warning("Key 不能为空")
                else:
                    k_up = k.upper()
                    v_up = v.upper()
                    new_map = dict(current_map)
                    new_map[k_up] = v_up
                    config.save_mapping_override(new_map)
                    st.session_state["notify"] = {"type": "add", "message": f"已添加/更新条目：{k_up} -> {v_up}"}
                    st.rerun()
            except Exception as e:
                st.error(f"添加失败: {e}")

        # 处理『删』
        if del_selected:
            try:
                merged_df = merge_view_to_full(full_df, edited_df)
                deleted_count = int(merged_df["delete"].sum()) if "delete" in merged_df.columns else 0
                kept = merged_df[~merged_df["delete"].astype(bool)].copy() if "delete" in merged_df.columns else merged_df
                new_map: dict[str, str] = {}
                for _, row in kept.iterrows():
                    k = str(row.get("key", "")).strip()
                    v = row.get("value", "")
                    if k:
                        k_up = k.upper()
                        v_str = "" if pd.isna(v) else str(v).strip()
                        new_map[k_up] = v_str.upper()
                config.save_mapping_override(new_map)
                st.session_state["notify"] = {"type": "delete", "message": f"已删除 {deleted_count} 条并保存。"}
                st.rerun()
            except Exception as e:
                st.error(f"删除失败: {e}")
        if save_mapping:
            try:
                merged_df = merge_view_to_full(full_df, edited_df)
                new_map: dict[str, str] = {}
                for _, row in merged_df.iterrows():
                    k = str(row.get("key", "")).strip()
                    v = row.get("value", "")
                    if k:
                        k_up = k.upper()
                        v_str = "" if pd.isna(v) else str(v).strip()
                        new_map[k_up] = v_str.upper()
                config.save_mapping_override(new_map)
                st.success(f"已保存到 mapping_override.json，并刷新内存映射，共 {len(config.missing_isin_or_stack_code_mapping_dict)} 条。")
                st.rerun()
            except Exception as e:
                st.error(f"保存失败: {e}")

    run_btn = st.button("Run Compare (Spectra↔HSBC)", disabled=not (spectra_file and hsbc_file), key="run_hsbc")
    run_vpfs_btn = st.button("Run Spectra↔VPFS", disabled=not (spectra_file and vpfs_file), key="run_vpfs")
    run_triple_btn = st.button("Run Triple Compare (Spectra↔HSBC↔VPFS)", disabled=not (spectra_file and hsbc_file and vpfs_file), key="run_triple")

    if run_btn and spectra_file and hsbc_file:
        spectra_bytes = BytesIO(spectra_file.read())
        hsbc_bytes = BytesIO(hsbc_file.read())
        with st.spinner("Comparing…"):
            result = run_compare_from_sources(spectra_bytes, hsbc_bytes)
        # 历史归档：保存输入与输出快照
        if getattr(config, "ENABLE_HISTORY", False):
            try:
                base_dir: Path = getattr(config, "HISTORY_DIR", Path(__file__).parent / "history")
                base_dir.mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir = base_dir / ts
                suffix = 1
                while out_dir.exists():
                    suffix += 1
                    out_dir = base_dir / f"{ts}_{suffix}"
                out_dir.mkdir(parents=False, exist_ok=False)

                # 写入输入快照（按原扩展名保留）
                spectra_ext = Path(spectra_file.name).suffix or ".xls"
                hsbc_ext = Path(hsbc_file.name).suffix or ".xlsx"
                (out_dir / f"spectra{spectra_ext}").write_bytes(spectra_bytes.getvalue())
                (out_dir / f"hsbc{hsbc_ext}").write_bytes(hsbc_bytes.getvalue())

                # 写入输出快照
                all_xlsx = result.get("all_sheets_xlsx")
                if isinstance(all_xlsx, (bytes, bytearray)):
                    (out_dir / "comparison_all.xlsx").write_bytes(all_xlsx)
            except Exception as e:
                st.warning(f"历史归档失败: {e}")
        st.session_state["result"] = result
        st.session_state["view"] = "results"
        st.rerun()
    # 三源对比
    if run_triple_btn and spectra_file and hsbc_file and vpfs_file:
        spectra_bytes = BytesIO(spectra_file.read())
        hsbc_bytes = BytesIO(hsbc_file.read())
        vpfs_bytes = BytesIO(vpfs_file.read())
        with st.spinner("Comparing Spectra↔HSBC↔VPFS…"):
            result = run_compare_triple_from_sources(spectra_bytes, hsbc_bytes, vpfs_bytes)
        if getattr(config, "ENABLE_HISTORY", False):
            try:
                base_dir: Path = getattr(config, "HISTORY_DIR", Path(__file__).parent / "history")
                base_dir.mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir = base_dir / ts
                suffix = 1
                while out_dir.exists():
                    suffix += 1
                    out_dir = base_dir / f"{ts}_{suffix}"
                out_dir.mkdir(parents=False, exist_ok=False)

                spectra_ext = Path(spectra_file.name).suffix or ".xls"
                hsbc_ext = Path(hsbc_file.name).suffix or ".xlsx"
                vpfs_ext = Path(vpfs_file.name).suffix or ".xls"
                (out_dir / f"spectra{spectra_ext}").write_bytes(spectra_bytes.getvalue())
                (out_dir / f"hsbc{hsbc_ext}").write_bytes(hsbc_bytes.getvalue())
                (out_dir / f"vpfs{vpfs_ext}").write_bytes(vpfs_bytes.getvalue())

                all_xlsx = result.get("all_sheets_xlsx")
                if isinstance(all_xlsx, (bytes, bytearray)):
                    (out_dir / "comparison_all.xlsx").write_bytes(all_xlsx)
            except Exception as e:
                st.warning(f"历史归档失败: {e}")
        st.session_state["result"] = result
        st.session_state["view"] = "results"
        st.rerun()
    # VPFS 跑批
    if run_vpfs_btn and spectra_file and vpfs_file:
        spectra_bytes = BytesIO(spectra_file.read())
        vpfs_bytes = BytesIO(vpfs_file.read())
        with st.spinner("Comparing Spectra↔VPFS…"):
            result = run_compare_spectra_vs_vpfs(spectra_bytes, vpfs_bytes)
        if getattr(config, "ENABLE_HISTORY", False):
            try:
                base_dir: Path = getattr(config, "HISTORY_DIR", Path(__file__).parent / "history")
                base_dir.mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir = base_dir / ts
                suffix = 1
                while out_dir.exists():
                    suffix += 1
                    out_dir = base_dir / f"{ts}_{suffix}"
                out_dir.mkdir(parents=False, exist_ok=False)

                spectra_ext = Path(spectra_file.name).suffix or ".xls"
                vpfs_ext = Path(vpfs_file.name).suffix or ".xls"
                (out_dir / f"spectra{spectra_ext}").write_bytes(spectra_bytes.getvalue())
                (out_dir / f"vpfs{vpfs_ext}").write_bytes(vpfs_bytes.getvalue())

                all_xlsx = result.get("all_sheets_xlsx")
                if isinstance(all_xlsx, (bytes, bytearray)):
                    (out_dir / "comparison_all.xlsx").write_bytes(all_xlsx)
            except Exception as e:
                st.warning(f"历史归档失败: {e}")
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

    # 将 diffs 放到第一个并默认显示
    tabs = st.tabs(["diffs", "comparison", "unmatched", "duplicates"]) 
    for tab_name, tab in zip(["diffs", "comparison", "unmatched", "duplicates"], tabs):
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
                    
                    # 检测是否为三源模式（存在 __equal_hsbc 或 __equal_vpfs 列）
                    is_triple_mode = any(f"{field}__equal_hsbc" in base.columns or f"{field}__equal_vpfs" in base.columns for field in SPECTRA_TO_HSBC_MAP.keys())
                    
                    if is_triple_mode:
                        # 三源模式：分别处理 HSBC 和 VPFS
                        for field in SPECTRA_TO_HSBC_MAP.keys():
                            left_alias = f"{field}__spectra"
                            hsbc_alias = f"{field}__hsbc"
                            vpfs_alias = f"{field}__vpfs"
                            hsbc_equal_col = f"{field}__equal_hsbc"
                            vpfs_equal_col = f"{field}__equal_vpfs"
                            
                            # HSBC 高亮
                            if hsbc_equal_col in base.columns:
                                cond_hsbc = ~base[hsbc_equal_col].fillna(False)
                                colors_hsbc = ["background-color: #ffff00" if v else "" for v in cond_hsbc.tolist()]
                                if left_alias in view.columns:
                                    styler = styler.apply(lambda s, c=colors_hsbc: c, subset=[left_alias])
                                if hsbc_alias in view.columns:
                                    styler = styler.apply(lambda s, c=colors_hsbc: c, subset=[hsbc_alias])
                            
                            # VPFS 高亮
                            if vpfs_equal_col in base.columns:
                                cond_vpfs = ~base[vpfs_equal_col].fillna(False)
                                colors_vpfs = ["background-color: #ffff00" if v else "" for v in cond_vpfs.tolist()]
                                if left_alias in view.columns:
                                    styler = styler.apply(lambda s, c=colors_vpfs: c, subset=[left_alias])
                                if vpfs_alias in view.columns:
                                    styler = styler.apply(lambda s, c=colors_vpfs: c, subset=[vpfs_alias])
                    else:
                        # 双源模式：HSBC 或 VPFS
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


