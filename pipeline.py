"""
pipeline.py
================
HKEFLPL Pick List (PL) generator pipeline.

This is a 1:1 port of the original Excel VBA system (Module1 .. Module7)
into pandas. Each VBA module maps to one function here:

  Module1  ->  format_inventory()
  Module2  ->  format_customer()  + filter_customer()      (slicer replacement)
  Module3  ->  filter_inventory_by_load_id()
  Module4  ->  compare_cus_vs_inv()
  Module5  ->  generate_pl()
  Module6  ->  fill_pl_blanks()
  Module7  ->  (handled in the app: archive current PL + reset session)

All functions are pure: they take DataFrames in and return DataFrames out,
so they are easy to unit test and reuse outside Streamlit.
"""

from __future__ import annotations

import io
import re
from datetime import datetime

import pandas as pd


# ---------------------------------------------------------------------------
# PL template column order (A .. W = 23 columns), reconstructed from
# Module5 write order + Module6 Find() header names.
# ---------------------------------------------------------------------------
PL_COLUMNS = [
    "CTN No",                # A
    "NO OF CTNs",            # B
    "UOM",                   # C
    "Ref No.",               # D  <- Pallet
    "Item",                  # E  (blank)
    "SL OR NON SL",          # F  <- user input
    "Order Ref",             # G  <- SO Number (Cus)
    "Composition",           # H  <- Supplier Desc (Inv)
    "SO Cust Order Number",  # I  <- Cus
    "Customer",              # J  <- Cust Name (Cus)
    "SKU BARCODE",           # K  <- Supplier (Inv)
    "Prod Code",             # L  <- Lot Number (Inv)
    "SO & PO",               # M  <- Plant split by ~ (left)
    "Season",                # N  <- Plant split by ~ (right)
    "Article No",            # O  <- Style (Inv)
    "Description",           # P  <- Color (Inv)
    "SIZE",                  # Q  <- Size (Inv)
    "QTY",                   # R  <- Actual Qty (Inv)
    "UOM ",                  # S  (second UOM, trailing space keeps it unique)
    "Net WT in Kgs",         # T  <- Current Net Weight (Inv), 3 decimals
    "GROSS WT in Kgs",       # U  <- Current Gross Weight (Inv), 3 decimals
    "Item Id",               # V  <- Inv
    "Vendor Name",           # W  <- Inv
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def find_col(df: pd.DataFrame, *candidates: str) -> str | None:
    """Case-insensitive, alias-aware column lookup.

    Tries exact (case-insensitive) match first, then a 'contains' fuzzy
    match. Returns the actual column name in df, or None.
    """
    cols = {str(c).strip().lower(): c for c in df.columns}
    # exact
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    # fuzzy: column name CONTAINS the full candidate phrase.
    # (We deliberately do NOT match the reverse direction — that would let a
    #  short generic column like "Supplier" satisfy a query for "Supplier Desc".)
    for cand in candidates:
        key = cand.strip().lower()
        if not key:
            continue
        for lc, orig in cols.items():
            if key in lc:
                return orig
    return None


def clean_number(val) -> str:
    """Replicates the VBA '.00 removal' / number-format '0'.

    1234.0 -> '1234', '1234.00' -> '1234', keeps real decimals otherwise,
    and leaves non-numeric text untouched.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s == "":
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except (ValueError, TypeError):
        # strip a trailing ".00" from text (VBA Replace ".00" -> "")
        return s.replace(".00", "")


def clean_barcode(val) -> str:
    """Extract the leading EAN/barcode digits.

    Handles values like '657001294220-GRN-1794391' (Outbound report)
    -> '657001294220', as well as plain '657001359806' and floats.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    m = re.match(r"(\d{4,})", s)
    if m:
        return m.group(1)
    return clean_number(s)


def _norm_key(val) -> str:
    """Normalize a barcode/UPC/Supplier value so '123', '123.0', ' 123 ',
    and '123-GRN-456' all compare equal in the comparison step."""
    return clean_barcode(val).strip()


# ---------------------------------------------------------------------------
# Module1 -> format_inventory
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Enrichment: pull blank Composition / Weights / Vendor from a SEPARATE
# inventory file, matched on a key column common to both files.
# (Use case: the Outbound Staged Licence Plate report has no Supplier Desc /
#  weights / Vendor Name; a full Inventory export does.)
# ---------------------------------------------------------------------------
ENRICH_FIELDS = {
    "Supplier Desc": ["Supplier Desc"],
    "Current Net Weight": ["Current Net Weight", "Net WT in Kgs"],
    "Current Gross Weight": ["Current Gross Weight", "GROSS WT in Kgs"],
    "Vendor Name": ["Vendor Name"],
}


def _key_series(df: pd.DataFrame, col: str, barcode: bool) -> list[str]:
    if barcode:
        return [clean_barcode(v) for v in df[col].tolist()]
    out = []
    for v in df[col].tolist():
        out.append("" if (v is None or (isinstance(v, float) and pd.isna(v)))
                   else str(v).strip())
    return out


def common_key_candidates(target_df: pd.DataFrame,
                          inv_df: pd.DataFrame) -> list[str]:
    """Columns present in BOTH files, ordered by how many values overlap
    (best join key first)."""
    pref = ["Stored Attribute Id", "Supplier", "Lot Number", "Item Id",
            "Style", "Pallet", "Hu Id", "Customer Ref Number", "Client So"]
    scored = []
    for name in pref:
        oc = find_col(target_df, name)
        ic = find_col(inv_df, name)
        if oc is None or ic is None:
            continue
        use_bc = name.strip().lower() == "supplier"
        ov = set(_key_series(target_df, oc, use_bc)) - {""}
        iv = set(_key_series(inv_df, ic, use_bc)) - {""}
        scored.append((len(ov & iv), name))
    scored.sort(key=lambda x: -x[0])
    return [n for _, n in scored]


def enrich_from_inventory(target_df: pd.DataFrame, inv_df: pd.DataFrame,
                          key: str = "Supplier") -> tuple[pd.DataFrame, dict]:
    """Fill Supplier Desc / Net&Gross Weight / Vendor Name into target_df by
    matching `key` against the inventory file. Only blank cells are filled.

    Returns (enriched_df, info) where info has key/fields/filled_rows/
    matched_keys.
    """
    out = target_df.copy()
    out_key_col = find_col(out, key)
    inv_key_col = find_col(inv_df, key)
    info = {"key": key, "fields": [], "filled_rows": 0, "matched_keys": 0}
    if out_key_col is None or inv_key_col is None:
        return out, info

    use_bc = key.strip().lower() == "supplier"

    src_cols = {}
    for target, aliases in ENRICH_FIELDS.items():
        c = find_col(inv_df, *aliases)
        if c is not None:
            src_cols[target] = c
    info["fields"] = list(src_cols.keys())
    if not src_cols:
        return out, info

    inv_keys = _key_series(inv_df, inv_key_col, use_bc)
    lookup: dict[str, dict] = {}
    for i, kv in enumerate(inv_keys):
        if not kv or kv in lookup:
            continue
        row = inv_df.iloc[i]
        lookup[kv] = {t: row.get(c, "") for t, c in src_cols.items()}
    info["matched_keys"] = len(lookup)

    for target in src_cols:
        if find_col(out, target) is None:
            out[target] = ""

    def _blank(v):
        return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""

    out_keys = _key_series(out, out_key_col, use_bc)
    filled = 0
    for idx, kv in zip(out.index, out_keys):
        if kv not in lookup:
            continue
        touched = False
        for target, val in lookup[kv].items():
            tcol = find_col(out, target)
            if _blank(out.at[idx, tcol]) and not _blank(val):
                out.at[idx, tcol] = val
                touched = True
        if touched:
            filled += 1
    info["filled_rows"] = filled
    return out, info


# ---------------------------------------------------------------------------
# Smart Excel reading: auto-detect the header row, keep meaningful-but-empty
# columns, drop only blank/Unnamed header columns.
# ---------------------------------------------------------------------------
INVENTORY_MARKERS = {"pallet", "actual qty", "supplier", "supplier desc",
                     "style", "plant", "lot number", "current net weight",
                     "hu id", "order number", "picked location"}
CUSTOMER_MARKERS = {"product upc", "so number", "pick qty", "country name",
                    "cust name", "so cust order number"}


def _detect_header_row(raw: pd.DataFrame, markers: set[str], max_scan: int = 6) -> int:
    """Pick the row (within the first few) that best looks like a header by
    counting how many known marker names it contains."""
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        vals = {str(v).strip().lower() for v in raw.iloc[i].tolist() if v is not None}
        score = len(vals & markers)
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def _read_table(file, markers: set[str], header_row: int | None = None) -> pd.DataFrame:
    raw = pd.read_excel(file, header=None, dtype=object)
    if header_row is None:
        header_row = _detect_header_row(raw, markers)
    cols = [str(c).strip() if c is not None else "" for c in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = cols
    # keep only columns whose HEADER is meaningful (drop blank / 'nan' / Unnamed)
    keep = [c for c in df.columns
            if c and str(c).strip().lower() not in ("nan", "none")
            and not str(c).startswith("Unnamed")]
    # de-duplicate while preserving order
    seen, keep_unique = set(), []
    for c in keep:
        if c not in seen:
            seen.add(c)
            keep_unique.append(c)
    df = df.loc[:, keep_unique]
    # drop fully-empty trailing data rows
    df = df.dropna(axis=0, how="all").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Module1 -> format_inventory
# ---------------------------------------------------------------------------
def read_inventory(file, skip_first_row: bool | None = None,
                   header_row: int | None = None) -> pd.DataFrame:
    """Read the inventory workbook with automatic header-row detection.

    - header_row=None (default): auto-detect which of the first rows is the
      real header (handles both the old 'title row on top' export and the
      new 'header on row 1' export).
    - skip_first_row: legacy convenience. True -> header_row=1, False -> 0.
      Only used when header_row is not given.
    """
    if header_row is None and skip_first_row is not None:
        header_row = 1 if skip_first_row else 0
    return _read_table(file, INVENTORY_MARKERS, header_row)


def format_inventory(df: pd.DataFrame) -> pd.DataFrame:
    """Module1: convert numeric columns to clean numbers / strip .00."""
    df = df.copy()
    for name in ["Pallet", "Current Gross Weight", "Current Net Weight",
                 "Actual Qty", "Supplier"]:
        col = find_col(df, name)
        if col is not None:
            df[col] = df[col].map(clean_number)
    return df


# ---------------------------------------------------------------------------
# Module2 -> format_customer + filter_customer (slicer replacement)
# ---------------------------------------------------------------------------
def read_customer(file, header_row: int | None = None) -> pd.DataFrame:
    return _read_table(file, CUSTOMER_MARKERS, header_row)


def format_customer(df: pd.DataFrame) -> pd.DataFrame:
    """Module2 formatting: Product UPC -> integer text, SO/QTY -> numeric."""
    df = df.copy()
    upc = find_col(df, "Product UPC")
    if upc is not None:
        df[upc] = df[upc].map(clean_number)
    for name in ["SO Number", "PICK QTY"]:
        col = find_col(df, name)
        if col is not None:
            df[col] = df[col].map(clean_number)
    return df


def customer_filter_options(df: pd.DataFrame) -> dict:
    """Returns available distinct values for the 'slicer' columns."""
    out = {}
    for name in ["SO Number", "Country Name", "PICK QTY"]:
        col = find_col(df, name)
        if col is not None:
            vals = sorted({str(v) for v in df[col].dropna().tolist()})
            out[name] = (col, vals)
    return out


def filter_customer(df: pd.DataFrame, selections: dict) -> pd.DataFrame:
    """Apply the slicer selections (multiselect) to the customer df.

    selections = {"SO Number": [..], "Country Name": [..], ...}
    Empty / missing selection = no filter on that column.
    """
    out = df.copy()
    for name, chosen in selections.items():
        if not chosen:
            continue
        col = find_col(out, name)
        if col is None:
            continue
        chosen_str = {str(c) for c in chosen}
        out = out[out[col].astype(str).isin(chosen_str)]
    return out


# ---------------------------------------------------------------------------
# Module3 -> filter_inventory_by_load_id
# ---------------------------------------------------------------------------
def filter_inventory_by_load_id(df: pd.DataFrame, load_ids: list[str],
                                load_col: str | None = None) -> pd.DataFrame:
    """Module3: wildcard (partial, case-insensitive) match on the Load ID
    column. VBA filtered Column C (the 3rd column); we try a named
    'Load ID' column first, then fall back to the 3rd column.
    """
    ids = [str(x).strip() for x in load_ids if str(x).strip()]
    if not ids:
        return df.iloc[0:0].copy()

    if load_col is None:
        load_col = find_col(df, "Load Id", "Load ID", "LoadID")
    if load_col is None:
        load_col = df.columns[2] if len(df.columns) >= 3 else df.columns[0]

    pattern = "|".join(re.escape(i) for i in ids)
    mask = df[load_col].astype(str).str.contains(pattern, case=False, na=False)
    return df[mask].copy()


# ---------------------------------------------------------------------------
# Module4 -> compare_cus_vs_inv
# ---------------------------------------------------------------------------
def compare_cus_vs_inv(cus_df: pd.DataFrame, inv_df: pd.DataFrame) -> pd.DataFrame:
    """Module4: aggregate Cus (UPC -> sum PICK QTY) vs Inv (Supplier ->
    sum Actual Qty), join on the identifier and flag MATCHED / MISMATCH.
    """
    upc = find_col(cus_df, "Product UPC")
    pick = find_col(cus_df, "PICK QTY")
    sup = find_col(inv_df, "Supplier")
    act = find_col(inv_df, "Actual Qty")

    missing = [n for n, c in [("Product UPC", upc), ("PICK QTY", pick),
                              ("Supplier", sup), ("Actual Qty", act)] if c is None]
    if missing:
        raise ValueError("Comparison සඳහා columns හමු නොවුණා: " + ", ".join(missing))

    def to_num(s):
        return pd.to_numeric(s.map(clean_number), errors="coerce").fillna(0)

    cus = cus_df.assign(_k=cus_df[upc].map(_norm_key), _q=to_num(cus_df[pick]))
    inv = inv_df.assign(_k=inv_df[sup].map(_norm_key), _q=to_num(inv_df[act]))

    cus_g = cus.groupby("_k")["_q"].sum()
    inv_g = inv.groupby("_k")["_q"].sum()

    rows = []
    for key, iqty in inv_g.items():
        cqty = float(cus_g.get(key, 0))
        iqty = float(iqty)
        status = "MATCHED" if cqty == iqty else "MISMATCH"
        rows.append({"ID": key, "Cus Qty": cqty, "Inv Qty": iqty, "Status": status})

    out = pd.DataFrame(rows, columns=["ID", "Cus Qty", "Inv Qty", "Status"])
    return out.sort_values(["Status", "ID"], ascending=[True, True]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Module5 -> generate_pl
# ---------------------------------------------------------------------------
def _split_plant(value):
    s = "" if value is None else str(value)
    if "~" in s:
        left, right = s.split("~", 1)
        return left, right
    return s, ""


def generate_pl(inv_df: pd.DataFrame, cus_df: pd.DataFrame, sl_flag: str) -> pd.DataFrame:
    """Module5: build the PL rows from filtered inventory + customer.

    Faithful to the VBA: the customer-order header fields (Order Ref,
    SO Cust Order Number, Customer) are taken from the FIRST customer row
    and applied to every PL line. Item details come per inventory row.
    """
    if cus_df.empty:
        cus_first = {}
    else:
        cus_first = cus_df.iloc[0].to_dict()

    def cval(*names):
        for n in names:
            col = find_col(cus_df, n)
            if col is not None:
                return cus_first.get(col, "")
        return ""

    so_number = clean_number(cval("SO Number"))
    so_cust_order_cus = cval("SO Cust Order Number")
    cust_name = cval("Cust Name", "Customer")

    rows = []
    for seq, (_, r) in enumerate(inv_df.iterrows(), start=1):
        def iv(*names):
            for n in names:
                col = find_col(inv_df, n)
                if col is not None:
                    v = r.get(col, "")
                    return "" if (v is None or (isinstance(v, float) and pd.isna(v))) else v
            return ""

        so_po, season = _split_plant(iv("Plant"))
        net = iv("Current Net Weight", "Net WT in Kgs")
        gross = iv("Current Gross Weight", "GROSS WT in Kgs")

        # "Cust Po Number වෙනුවට Order Number" — customer order number column
        # now comes from the report's 'Order Number' when present.
        order_number = iv("Order Number")
        so_cust_order = order_number if str(order_number).strip() else so_cust_order_cus

        rows.append({
            "CTN No": seq,
            "NO OF CTNs": 1,
            "UOM": "CTN",
            "Ref No.": iv("Pallet", "Hu Id", "Licence Plate", "License Plate"),
            "Item": "",
            "SL OR NON SL": sl_flag,
            "Order Ref": so_number if str(so_number).strip() else iv("Client So", "Order Number"),
            "Composition": iv("Supplier Desc"),
            "SO Cust Order Number": so_cust_order,
            "Customer": cust_name,
            "SKU BARCODE": clean_barcode(iv("Supplier")),
            "Prod Code": iv("Lot Number"),
            "SO & PO": so_po,
            "Season": season,
            "Article No": iv("Style"),
            "Description": iv("Color"),
            "SIZE": iv("Size"),
            "QTY": iv("Actual Qty"),
            "UOM ": "PCS",
            "Net WT in Kgs": _fmt3(net) if str(net).strip() else "",
            "GROSS WT in Kgs": _fmt3(gross) if str(gross).strip() else "",
            "Item Id": iv("Item Id"),
            "Vendor Name": iv("Vendor Name"),
        })

    return pd.DataFrame(rows, columns=PL_COLUMNS)


def _fmt3(value) -> str:
    try:
        return f"{float(clean_number(value)):.3f}"
    except (ValueError, TypeError):
        return ""


# ---------------------------------------------------------------------------
# Module6 -> fill_pl_blanks
# ---------------------------------------------------------------------------
def fill_pl_blanks(pl_df: pd.DataFrame, inv_df: pd.DataFrame) -> pd.DataFrame:
    """Module6: for PL rows with a blank Composition, find an inventory
    row whose Pallet is a substring of the PL Ref No. (partial match) and
    fill Composition / Net WT / Gross WT / Vendor Name from it.

    First drops inventory rows with a blank 'Supplier Desc' (Inventory2).
    """
    pl = pl_df.copy()
    sup_desc = find_col(inv_df, "Supplier Desc")
    pallet = find_col(inv_df, "Pallet")
    if sup_desc is None or pallet is None:
        return pl

    def _blank(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return True
        return str(v).strip() == ""

    # Inventory2: drop rows whose Supplier Desc is blank (VBA Module6 step 1)
    inv2 = inv_df[~inv_df[sup_desc].map(_blank)].copy()
    inv_net = find_col(inv2, "Net WT in Kgs", "Current Net Weight")
    inv_gross = find_col(inv2, "GROSS WT in Kgs", "Current Gross Weight")
    inv_vendor = find_col(inv2, "Vendor Name")

    lookups = []
    for _, ir in inv2.iterrows():
        pid = "" if _blank(ir.get(pallet)) else str(ir.get(pallet)).strip()
        if pid:
            lookups.append((pid, ir))

    for idx, prow in pl.iterrows():
        if not _blank(prow.get("Composition")):
            continue
        ref = str(prow.get("Ref No.", ""))
        for pid, ir in lookups:
            if pid.lower() in ref.lower():
                pl.at[idx, "Composition"] = ir.get(sup_desc, "")
                if inv_net:
                    pl.at[idx, "Net WT in Kgs"] = _fmt3(ir.get(inv_net, ""))
                if inv_gross:
                    pl.at[idx, "GROSS WT in Kgs"] = _fmt3(ir.get(inv_gross, ""))
                if inv_vendor:
                    pl.at[idx, "Vendor Name"] = ir.get(inv_vendor, "")
                break
    return pl


# ---------------------------------------------------------------------------
# Excel export (PL.xlsx) and Comparison summary export
# ---------------------------------------------------------------------------
def pl_to_xlsx_bytes(pl_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        pl_df.to_excel(xl, index=False, sheet_name="PL")
    buf.seek(0)
    return buf.read()


def summary_to_xlsx_bytes(summary_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        summary_df.to_excel(xl, index=False, sheet_name="Comparison_Summary")
    buf.seek(0)
    return buf.read()


def new_run_id() -> str:
    return datetime.now().strftime("RUN-%Y%m%d-%H%M%S")
