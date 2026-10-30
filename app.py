#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py
======

Streamlit front-end for the ERP activity merge.

    streamlit run app.py

Workflow for the user:
    1. sign in with their own ERP account,
    2. upload ``mel raw data.xlsx`` (with Sheet2 holding the project list),
    3. press one button and download the same workbook with Sheet3 filled in.

The uploaded file is processed in a temporary folder and never modified in
place; the user gets a download button for the result.
"""

from __future__ import annotations

import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

try:  # module name in the development workspace
    from erp_export_merge import BASE_URL, AuthError, ErpError
except ImportError:  # deployed next to the CLI copy named erp_export.py
    from erp_export import BASE_URL, AuthError, ErpError  # type: ignore[no-redef]

from merge_pipeline import run_merge

st.set_page_config(
    page_title="HERD M&E · ERP Activity Merge",
    page_icon="📊",
    layout="wide",
)


def _reset() -> None:
    for key in ("result", "last_upload"):
        st.session_state.pop(key, None)


def _candidate_dirs():
    """Writable folders for the uploaded copy: system temp, then ./herd_work."""
    try:
        yield Path(tempfile.mkdtemp(prefix="herd_erp_"))
    except OSError:
        pass
    fallback = Path.cwd() / "herd_work" / datetime.now().strftime("%Y%m%d_%H%M%S")
    fallback.mkdir(parents=True, exist_ok=True)
    yield fallback


def _save_upload(uploaded) -> Path:
    data = uploaded.getvalue()
    last_error: OSError | None = None
    for folder in _candidate_dirs():
        try:
            path = folder / uploaded.name
            path.write_bytes(data)
            return path
        except OSError as exc:
            last_error = exc
    raise last_error or OSError("no writable folder")


def main() -> None:
    st.title("HERD M&E · ERP Activity Merge")
    st.caption(
        "Upload your M&E workbook, sign in to the ERP, and download the same file "
        "with every project's activities merged into Sheet3."
    )

    with st.sidebar:
        st.header("1 · ERP login")
        username = st.text_input("ERP username", placeholder="name@herdint.com")
        password = st.text_input("ERP password", type="password")
        st.caption("Your password is used only for this run. It is not saved or logged.")

        with st.expander("Advanced settings"):
            base_url = st.text_input("ERP address", value=BASE_URL)
            sheet = st.text_input("Sheet with the project list", value="Sheet2")
            out_sheet = st.text_input("Sheet to write the result", value="Sheet3")

        st.divider()
        if st.button("Clear results"):
            _reset()
            st.rerun()

    st.header("2 · Your workbook")
    uploaded = st.file_uploader(
        "Choose the Excel file (must contain the project list sheet)",
        type=["xlsx"],
    )
    if uploaded is not None and uploaded.name != st.session_state.get("last_upload"):
        st.session_state["last_upload"] = uploaded.name
        st.session_state.pop("result", None)

    st.header("3 · Run")
    ready = bool(uploaded is not None and username and password)
    if not ready:
        st.info("Enter your ERP login and choose a file to enable the button.")

    start = st.button(
        "Download activities and merge",
        type="primary",
        disabled=not ready,
        use_container_width=False,
    )

    if start and uploaded is not None:
        try:
            source = _save_upload(uploaded)
        except OSError as exc:
            st.error(f"Could not save the uploaded file to a working folder: {exc}")
            return
        workdir = source.parent

        progress_bar = st.progress(0.0)
        status = st.empty()

        def on_progress(done: int, total: int, message: str) -> None:
            progress_bar.progress(min(done / max(total, 1), 1.0))
            status.info(message)

        try:
            result = run_merge(
                workbook_path=source,
                username=username,
                password=password,
                sheet=sheet,
                out_sheet=out_sheet,
                base_url=base_url,
                exports_dir=workdir / "erp_exports",
                progress=on_progress,
            )
        except AuthError:
            progress_bar.empty()
            status.empty()
            st.error(
                "Login failed. Check your ERP username and password, then try again."
            )
            return
        except ErpError as exc:
            progress_bar.empty()
            status.empty()
            st.error(f"The ERP returned an error: {exc}")
            return
        except Exception as exc:  # noqa: BLE001
            progress_bar.empty()
            status.empty()
            st.error("Something went wrong while building the file.")
            with st.expander("Technical details"):
                st.code(f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}")
            return

        progress_bar.progress(1.0)
        status.success("Finished.")

        st.session_state["result"] = {
            "data": source.read_bytes(),
            "name": uploaded.name,
            "rows": result.merged_rows,
            "ok": result.ok_projects,
            "total": result.total_projects,
            "mapping": result.mapping,
            "warnings": result.warnings,
        }

    result = st.session_state.get("result")
    if not result:
        return

    st.divider()
    missing = result["total"] - result["ok"]
    if missing:
        st.warning(
            f"{result['ok']} of {result['total']} projects were merged; "
            f"{missing} need attention (see below)."
        )
    else:
        st.success(
            f"All {result['total']} projects merged — {result['rows']:,} activity rows."
        )

    st.download_button(
        "⬇ Download the updated workbook",
        data=result["data"],
        file_name=result["name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(
        f"The workbook now has a '{out_sheet}' sheet with the merged activities "
        "and a 'Project Map' sheet showing how each project was matched."
    )

    for warning in result["warnings"][:15]:
        st.write(f"- {warning}")

    with st.expander("Project matching details", expanded=bool(missing)):
        frame = pd.DataFrame(result["mapping"])
        if not frame.empty:
            st.dataframe(frame, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
