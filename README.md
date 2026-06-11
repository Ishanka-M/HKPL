# HKEFLPL Pick List Generator — Streamlit (stateless)

පරණ Excel/VBA system එක (Module1–Module7) Streamlit web app එකකට ගේපු එක.
**Data save වෙන්නේ නෑ** — DB/GitHub/history නෑ. හැම session එකක්ම files
browser එකේ memory එකේ process කරලා **PL.xlsx download** කරනවා.
`D:\HKEFLPL\...` paths අවශ්‍ය නෑ.

## VBA → Python mapping

| පරණ VBA | මෙහි |
|---|---|
| Module1 (FormatInventoryFile) | `pipeline.format_inventory` |
| Module2 (slicer form + filter) | `format_customer` + `filter_customer` (Streamlit multiselect = slicer) |
| Module3 (Load ID filter) | `filter_inventory_by_load_id` |
| Module4 (compare Cus vs Inv) | `compare_cus_vs_inv` |
| Module5 (Generate PL, 23 cols) | `generate_pl` |
| Module6 (blank fill, partial match) | `fill_pl_blanks` |
| Module7 (reset) | sidebar "Clear all (session reset)" |

## Files
```
app.py             Streamlit UI (stateless pipeline)
pipeline.py        Module1-7 logic (pandas) — pure functions
requirements.txt
.streamlit/config.toml          dark theme
.streamlit/secrets.toml.example optional APP_PASSWORD only
```

## Run

```bash
python -m venv .venv
# Windows:  .venv\Scripts\activate     | mac/linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

> Microsoft-Store Python වල PATH issues නම් `py -m venv .venv` /
> `py -m pip ...` පාවිච්චි කරන්න.

Secrets අවශ්‍ය නෑ. Shared app එකකට password එකක් ඕන නම් විතරක්,
`.streamlit/secrets.toml.example` -> `secrets.toml` rename කරලා
`APP_PASSWORD` දාන්න.

## Flow

1. **Source report** upload (Outbound Staged / Inventory). Header
   auto-detect. ඕන නම් **Enrichment Inventory** එකකුත් දාන්න.
2. **CustR** (optional) upload + slicer filter.
3. **Order / Load ID** filter (column selectable).
4. **Compare** Customer vs Source (live, download xlsx).
5. **Generate PL** (SL/NON SL) -> **download PL.xlsx**.

## Supported source formats (header auto-detect + alias matching)

1. **Inventory Report** — `Pallet`, `Supplier Desc` (composition),
   `Current Net/Gross Weight`, `Vendor Name` තියෙනවා.
2. **Outbound Staged Licence Plate** — `Pallet` වෙනුවට **`Hu Id`** ->
   Ref No.; **`Order Number`** -> SO Cust Order Number (Cust PO වෙනුවට);
   `Supplier` GRN suffix barcode එකට clean වෙනවා;
   `Supplier Desc`/weights/`Vendor Name` නැති නිසා blank.

### Enrichment — blank data වෙනම Inventory එකකින් පුරවන්න

Step 1 එකේ **Enrichment Inventory (optional)** upload කරලා, Outbound
එකේ හිස් `Supplier Desc` / `Net&Gross Weight` / `Vendor Name` පුරවගන්න:

- **Match key**: දෙ-file එකේම තියෙන column (default `Supplier` = barcode;
  value overlap වැඩිම එක මුලින් suggest වෙනවා).
- Barcode GRN suffix normalize වෙලා match වෙනවා
  (`657001294220-GRN-1794391` = `657001294220`).
- හිස් cells විතරක් පිරෙනවා.

## PL column mapping (Module5)

`CTN No, NO OF CTNs(=1), UOM(=CTN), Ref No.(Pallet/Hu Id), Item, SL OR
NON SL, Order Ref(SO Number/Client So), Composition(Supplier Desc),
SO Cust Order Number(Order Number), Customer(Cust Name), SKU
BARCODE(Supplier barcode), Prod Code(Lot Number), SO & PO + Season
(Plant split by ~), Article No(Style), Description(Color), SIZE(Size),
QTY(Actual Qty), UOM(=PCS), Net WT in Kgs, GROSS WT in Kgs, Item Id,
Vendor Name`.
