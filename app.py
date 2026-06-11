r"""
app.py — HKEFLPL Pick List Generator (Streamlit, stateless)
===========================================================
Replaces the Excel/VBA Module1..Module7 workflow with a web app:

  Upload Source report (+ optional Enrichment inventory)
  -> format -> slicer-filter customer -> filter by Order/Load ID
  -> compare -> generate PL -> fill blanks -> download PL.xlsx

No database, no GitHub, no run-history. Nothing is saved server-side:
each session processes files in memory and you download the result.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

import pipeline as P

st.set_page_config(page_title="HKEFLPL Pick List Generator",
                   page_icon="📦", layout="wide")

SS = st.session_state
for k, default in {
    "inv_fmt": None, "inv_filtered": None,
    "cus_fmt": None, "cus_filtered": None,
    "comparison": None, "pl": None, "authed": False,
}.items():
    SS.setdefault(k, default)


def df_show(df, **kw):
    st.dataframe(df, use_container_width=True, hide_index=True, **kw)


XLSX_MIME = ("application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet")

# --------------------------------------------------------------------------
# Optional password gate (only if APP_PASSWORD secret is set)
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("📦 HKEFLPL")
    st.caption("Pick List generator — data save වෙන්නේ නෑ. "
               "Files browser එකේ memory එකේ process වෙලා, PL.xlsx download වෙනවා.")
    try:
        app_pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        app_pw = None
    if app_pw and not SS.authed:
        pw = st.text_input("App password", type="password")
        if pw:
            SS.authed = (pw == app_pw)
            if not SS.authed:
                st.error("Password වැරදියි.")
    elif app_pw:
        st.success("Unlocked")
    st.divider()
    if st.button("🧹 Clear all (session reset)"):
        for k in ["inv_fmt", "inv_filtered", "cus_fmt", "cus_filtered",
                  "comparison", "pl"]:
            SS[k] = None
        st.rerun()

if app_pw and not SS.authed:
    st.info("Sidebar එකේ app password එක දාලා unlock කරන්න.")
    st.stop()


# ==========================================================================
# PIPELINE
# ==========================================================================
st.subheader("Pick List generate කරන steps")
load_ids: list[str] = []
sl_flag = "SL"

# ---- Step 1: Source report (+ optional enrichment inventory) ----
with st.expander("1️⃣ Source report upload + format", expanded=True):
    inv_file = st.file_uploader(
        "Main report — Outbound Staged / Inventory (.xlsx)",
        type=["xlsx", "xls"], key="inv_up")
    c1, c2 = st.columns([1, 2])
    auto_hdr = c1.checkbox("Header auto-detect", value=True,
                           help="මුල් rows කිහිපය scan කරලා header row එක "
                                "හොයාගන්නවා (title row තිබුණත් නැතත් වැඩ).")
    manual_row = 0
    if not auto_hdr:
        manual_row = c2.number_input("Header row (0 = මුල් row)",
                                     min_value=0, max_value=10, value=0, step=1)

    enrich_file = st.file_uploader(
        "Enrichment Inventory (optional) — Composition / Net&Gross Weight / "
        "Vendor Name පුරවන්න", type=["xlsx", "xls"], key="enrich_up")

    if inv_file is not None:
        base = None
        try:
            base = P.format_inventory(P.read_inventory(
                inv_file, header_row=(None if auto_hdr else int(manual_row))))
        except Exception as e:
            st.error(f"Main report read error: {e}")

        info = None
        if base is not None and enrich_file is not None:
            try:
                enr = P.format_inventory(P.read_inventory(enrich_file))
                keys = P.common_key_candidates(base, enr)
                if keys:
                    key = st.selectbox(
                        "Match key (දෙ-file එකේම තියෙන column — overlap වැඩිම "
                        "එක මුලින්)", keys, key="enrich_key",
                        help="Composition/Weight/Vendor මේ key එකෙන් match "
                             "කරලා පුරවනවා. සාමාන්‍යයෙන් 'Supplier' (barcode).")
                    base, info = P.enrich_from_inventory(base, enr, key=key)
                else:
                    st.warning("දෙ-file එකේම common column එකක් හමු නොවුණා.")
            except Exception as e:
                st.error(f"Enrichment error: {e}")

        if base is not None:
            SS.inv_fmt = base
            msg = f"Loaded: {len(base)} rows, {len(base.columns)} cols"
            if info and info.get("fields"):
                msg += (f" | enriched {info['filled_rows']} rows via "
                        f"'{info['key']}' ({', '.join(info['fields'])})")
                if info["filled_rows"] == 0:
                    st.warning("Enrichment: row 0ක් filled — match key values "
                               "දෙ-file එකේ overlap වෙන්නේ නැද්ද බලන්න.")
            st.success(msg)
            df_show(base.head(20))

# ---- Step 2: Customer (CustR) format + slicer filter ----
with st.expander("2️⃣ Customer (CustR) upload + slicer filter"):
    cus_file = st.file_uploader("CustR.xlsx (optional)", type=["xlsx", "xls"],
                                key="cus_up")
    if cus_file is not None:
        try:
            SS.cus_fmt = P.format_customer(P.read_customer(cus_file))
        except Exception as e:
            st.error(f"CustR read error: {e}")

    if SS.cus_fmt is not None:
        opts = P.customer_filter_options(SS.cus_fmt)
        st.caption("Slicer වෙනුවට — හිස් තැබුවොත් filter නෑ:")
        selections = {}
        cols = st.columns(max(1, len(opts)))
        for (name, (col, vals)), cc in zip(opts.items(), cols):
            selections[name] = cc.multiselect(name, vals, key=f"sl_{name}")
        SS.cus_filtered = P.filter_customer(SS.cus_fmt, selections)
        st.success(f"Filtered customer rows: {len(SS.cus_filtered)}")
        df_show(SS.cus_filtered.head(20))

# ---- Step 3: Filter source by Order / Load ID ----
with st.expander("3️⃣ Source — Order / Load ID filter"):
    filt_col = None
    if SS.inv_fmt is not None:
        cols = list(SS.inv_fmt.columns)
        default = (P.find_col(SS.inv_fmt, "Order Number")
                   or P.find_col(SS.inv_fmt, "Load Id", "Load ID")
                   or P.find_col(SS.inv_fmt, "Po Number")
                   or P.find_col(SS.inv_fmt, "Customer Ref Number")
                   or cols[0])
        filt_col = st.selectbox(
            "Filter column", cols,
            index=cols.index(default) if default in cols else 0,
            help="Outbound report නම් සාමාන්‍යයෙන් 'Order Number'.")
    load_ids_raw = st.text_input(
        "ID(s) — කොමා වලින් වෙන් කරන්න",
        placeholder="HKEFL-SO-26277-S-006")
    load_ids = [x.strip() for x in load_ids_raw.split(",") if x.strip()]
    if st.button("Filter", disabled=SS.inv_fmt is None):
        if SS.inv_fmt is None:
            st.warning("මුලින්ම Step 1 කරන්න.")
        elif not load_ids:
            st.warning("ID එකක් දාන්න.")
        else:
            SS.inv_filtered = P.filter_inventory_by_load_id(
                SS.inv_fmt, load_ids, load_col=filt_col)
            st.success(f"Matched rows: {len(SS.inv_filtered)} "
                       f"(column: {filt_col})")
    if SS.inv_filtered is not None:
        df_show(SS.inv_filtered.head(20))

# ---- Step 4: Compare (Customer vs Source) ----
with st.expander("4️⃣ Compare — Customer vs Source"):
    src = st.radio("Customer source", ["Filtered", "Full"], horizontal=True)
    if st.button("Run comparison", disabled=SS.inv_fmt is None):
        cus_df = (SS.cus_filtered if (src == "Filtered" and SS.cus_filtered is not None)
                  else SS.cus_fmt)
        inv_df = SS.inv_filtered if SS.inv_filtered is not None else SS.inv_fmt
        if cus_df is None or inv_df is None:
            st.warning("Customer සහ Source දෙකම ඕන.")
        else:
            try:
                SS.comparison = P.compare_cus_vs_inv(cus_df, inv_df)
            except ValueError as e:
                st.error(str(e))
    if SS.comparison is not None:
        mism = int((SS.comparison["Status"] == "MISMATCH").sum())
        tot = len(SS.comparison)
        (st.error if mism else st.success)(f"Total: {tot} | Mismatches: {mism}")

        def hl(row):
            color = ("background-color:#7f1d1d;color:#fff"
                     if row["Status"] == "MISMATCH" else "")
            return [color] * len(row)
        st.dataframe(SS.comparison.style.apply(hl, axis=1),
                     use_container_width=True, hide_index=True)
        st.download_button("⬇️ Summary xlsx",
                           P.summary_to_xlsx_bytes(SS.comparison),
                           file_name="Comparison_Summary.xlsx", mime=XLSX_MIME)

# ---- Step 5 + 6: Generate PL + fill blanks ----
with st.expander("5️⃣ Generate Pick List (PL)", expanded=True):
    sl_flag = st.radio("SL OR NON SL", ["SL", "NON SL"], horizontal=True)
    do_fill = st.checkbox("blank Composition/WT/Vendor partial-match වලින් "
                          "පුරවන්න (Module6)", value=True)
    if st.button("🚀 Generate PL", type="primary",
                 disabled=SS.inv_filtered is None):
        if SS.inv_filtered is None:
            st.warning("Step 3 (filter) කරන්න.")
        else:
            cus_df = SS.cus_filtered if SS.cus_filtered is not None else SS.cus_fmt
            if cus_df is None:
                cus_df = pd.DataFrame()
            pl = P.generate_pl(SS.inv_filtered, cus_df, sl_flag)
            if do_fill:
                pl = P.fill_pl_blanks(pl, SS.inv_fmt)
            SS.pl = pl
            st.success(f"PL generated: {len(pl)} lines")
    if SS.pl is not None:
        df_show(SS.pl)
        st.download_button("⬇️ PL.xlsx", P.pl_to_xlsx_bytes(SS.pl.fillna("")),
                           file_name="PL.xlsx", mime=XLSX_MIME, type="primary")
