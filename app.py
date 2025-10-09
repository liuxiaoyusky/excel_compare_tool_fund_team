"""Streamlit front-end for the NAV validation pipeline."""
from __future__ import annotations

from io import BytesIO
from typing import Sequence

import pandas as pd
import streamlit as st

from nav_checker import (
    NavValidationContext,
    NavValidator,
    SpectraInboundRepository,
    ValidateNavUseCase,
    HsbcAuthoritativeRepository,
)
from nav_checker.domain.models import NavRecord
from nav_checker.domain.results import ValidationReport
from nav_checker.presentation.diff_report import discrepancies_to_rows, render_csv, render_html
from nav_checker.infrastructure.storage import mapping_store


st.set_page_config(page_title="NAV Validator", layout="wide")
st.title("NAV Validation Tool")


def load_mapping_dataframe() -> pd.DataFrame:
    mapping = mapping_store.load_mapping()
    return pd.DataFrame(
        [{"key": key, "value": value, "delete": False} for key, value in sorted(mapping.items())],
        columns=["key", "value", "delete"],
    )


def records_to_dataframe(records: Sequence[NavRecord]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "instrument_id": r.identity.instrument_id,
                "date": r.identity.nav_date,
                "currency": r.identity.currency,
                "share_class": r.identity.share_class,
                "nav": r.nav,
                "source": r.source,
                "file_hash": r.file_hash,
                "as_of": r.as_of,
                "lineage": r.lineage,
            }
            for r in records
        ]
    )


def run_validation(spectra_bytes: bytes, hsbc_bytes: bytes) -> tuple[ValidationReport, Sequence[NavRecord], Sequence[NavRecord]]:
    inbound_repo = SpectraInboundRepository(BytesIO(spectra_bytes))
    authoritative_repo = HsbcAuthoritativeRepository(BytesIO(hsbc_bytes))
    context = NavValidationContext(
        inbound_repository=inbound_repo,
        authoritative_repository=authoritative_repo,
        validator=NavValidator(),
    )
    use_case = ValidateNavUseCase(context)
    return use_case.execute()


if "view" not in st.session_state:
    st.session_state["view"] = "compare"
if "result" not in st.session_state:
    st.session_state["result"] = None


if st.session_state["view"] == "compare":
    col1, col2 = st.columns(2)
    with col1:
        spectra_file = st.file_uploader("Upload Spectra file", type=["xls", "xlsx"])
    with col2:
        hsbc_file = st.file_uploader("Upload HSBC file", type=["xls", "xlsx"])

    st.subheader("Security Mapping Editor")
    with st.expander("Manage missing ISIN / stack code mapping", expanded=True):
        mapping_df = load_mapping_dataframe()
        search_text = st.text_input("Search key/value", key="mapping_search")
        if search_text:
            pattern = str(search_text).strip().lower()
            mask = mapping_df.apply(
                lambda row: pattern in str(row.get("key", "")).lower()
                or pattern in str(row.get("value", "")).lower(),
                axis=1,
            )
            view_df = mapping_df[mask].copy()
        else:
            view_df = mapping_df.copy()
        st.caption(f"Total {len(mapping_df)} entries; showing {len(view_df)}")
        edited_df = st.data_editor(
            view_df,
            num_rows="dynamic",
            hide_index=True,
            key="mapping_editor",
            use_container_width=True,
        )

        col_add1, col_add2, col_add3 = st.columns([2, 2, 1])
        with col_add1:
            new_key = st.text_input("New key", key="new_map_key")
        with col_add2:
            new_val = st.text_input("New value", key="new_map_val")
        with col_add3:
            add_entry = st.button("Add", key="add_mapping_btn")

        col_ops1, col_ops2 = st.columns([1, 1])
        with col_ops1:
            del_selected = st.button("Delete selected", key="delete_selected_btn")
        with col_ops2:
            save_mapping = st.button("Save", key="save_mapping_btn")

        def merge_view(full_df: pd.DataFrame, view_df_edited: pd.DataFrame) -> pd.DataFrame:
            if not isinstance(view_df_edited, pd.DataFrame):
                return full_df.copy()
            view_df_edited = view_df_edited.copy()
            view_df_edited["key"] = view_df_edited["key"].astype(str).str.strip()
            view_df_edited = view_df_edited[view_df_edited["key"] != ""]
            view_keys = set(view_df_edited["key"].str.upper())
            rest = full_df[~full_df["key"].str.upper().isin(view_keys)].copy()
            merged = pd.concat([rest, view_df_edited], ignore_index=True)
            return merged

        if add_entry:
            key = (new_key or "").strip().upper()
            value = (new_val or "").strip().upper()
            if not key:
                st.warning("Key cannot be empty")
            else:
                merged = merge_view(mapping_df, edited_df)
                merged = pd.concat(
                    [merged, pd.DataFrame([{"key": key, "value": value, "delete": False}])],
                    ignore_index=True,
                )
                merged_clean = {row["key"].upper(): str(row["value"]).upper() for _, row in merged.iterrows() if row["key"]}
                mapping_store.save_mapping(merged_clean)
                st.success(f"Added {key} -> {value}")
                st.experimental_rerun()

        if del_selected:
            merged = merge_view(mapping_df, edited_df)
            if "delete" in merged.columns:
                merged = merged[~merged["delete"].astype(bool)]
            cleaned = {row["key"].upper(): str(row["value"]).upper() for _, row in merged.iterrows() if row["key"]}
            mapping_store.save_mapping(cleaned)
            st.success("Deleted selected entries")
            st.experimental_rerun()

        if save_mapping:
            merged = merge_view(mapping_df, edited_df)
            cleaned = {row["key"].upper(): str(row["value"]).upper() for _, row in merged.iterrows() if row["key"]}
            mapping_store.save_mapping(cleaned)
            st.success("Mapping saved")
            st.experimental_rerun()

    run_btn = st.button("Run Validation", disabled=not (spectra_file and hsbc_file))
    if run_btn and spectra_file and hsbc_file:
        spectra_bytes = spectra_file.read()
        hsbc_bytes = hsbc_file.read()
        with st.spinner("Validating..."):
            report, inbound_records, authoritative_records = run_validation(spectra_bytes, hsbc_bytes)
        st.session_state["result"] = {
            "report": report,
            "inbound": inbound_records,
            "authoritative": authoritative_records,
            "diff_csv": render_csv(tuple(report.iter_all_discrepancies())),
            "diff_html": render_html(report),
        }
        st.session_state["view"] = "results"
        st.experimental_rerun()
else:
    back_clicked = st.button("← Back", key="back_to_compare")
    if back_clicked:
        st.session_state["view"] = "compare"
        st.session_state["result"] = None
        st.experimental_rerun()

    result = st.session_state.get("result")
    if not result:
        st.info("No results available. Upload files and run validation first.")
    else:
        report: ValidationReport = result["report"]
        inbound_records: Sequence[NavRecord] = result["inbound"]
        authoritative_records: Sequence[NavRecord] = result["authoritative"]

        st.subheader("Summary")
        summary = report.summary
        st.metric("Inbound records", summary.total_inbound)
        st.metric("Authoritative records", summary.total_authoritative)
        st.metric("Wrong numbers", summary.wrong_numbers)
        st.metric("Missing in inbound", summary.missing_in_inbound)
        st.metric("Unauthorized inbound", summary.unauthorized_in_inbound)
        st.metric("Duplicates", summary.duplicates)

        tabs = st.tabs(["Diffs", "Inbound", "Authoritative"])
        with tabs[0]:
            discrepancies_df = pd.DataFrame(discrepancies_to_rows(tuple(report.iter_all_discrepancies())))
            st.dataframe(discrepancies_df)
            st.download_button(
                "Download diff CSV",
                data=result["diff_csv"],
                file_name="nav_diff.csv",
                mime="text/csv",
            )
            st.download_button(
                "Download diff HTML",
                data=result["diff_html"].encode("utf-8"),
                file_name="nav_diff.html",
                mime="text/html",
            )
        with tabs[1]:
            st.dataframe(records_to_dataframe(inbound_records))
        with tabs[2]:
            st.dataframe(records_to_dataframe(authoritative_records))
