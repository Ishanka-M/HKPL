# HKEFLPL Pick List Generator — Streamlit + GitHub JSON

පරණ Excel/VBA system එක (Module1–Module7) Streamlit web app එකකට ගේපු එක.
Database එකක් වෙනුවට — **GitHub repo එකක JSON file එකක්** "DB" විදිහට update වෙනවා.
`D:\HKEFLPL\...` paths අවශ්‍ය නෑ; files upload කරනවා.

## VBA → Python mapping

| පරණ VBA | මෙහි |
|---|---|
| Module1 (FormatInventoryFile) | `pipeline.format_inventory` |
| Module2 (slicer form + filter) | `format_customer` + `filter_customer` (Streamlit multiselect = slicer) |
| Module3 (Load ID filter) | `filter_inventory_by_load_id` |
| Module4 (compare Cus vs Inv) | `compare_cus_vs_inv` |
| Module5 (Generate PL, 23 cols) | `generate_pl` |
| Module6 (blank fill, partial match) | `fill_pl_blanks` |
| Module7 (reset PL + clear trash) | "Reset current" button + `reset_current` |

## Files
```
app.py             Streamlit UI (pipeline + History + Comparison tabs)
pipeline.py        Module1-7 logic (pandas) — pure functions
github_store.py    GitHub JSON read/write + optimistic locking (2-3 users)
requirements.txt
.streamlit/config.toml          dark theme
.streamlit/secrets.toml.example secrets template (rename + fill)
```

## 1. Local run

```bash
python -m venv .venv
# Windows:  .venv\Scripts\activate     | mac/linux: source .venv/bin/activate
pip install -r requirements.txt

# secrets template එක copy කරලා values දාන්න
copy .streamlit\secrets.toml.example .streamlit\secrets.toml   # Windows
# cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # mac/linux

streamlit run app.py
```

> **සටහන:** Microsoft-Store Python වල PATH issues තියෙනවා නම්,
> `py -m venv .venv` / `py -m pip ...` පාවිච්චි කරන්න.

## 2. GitHub setup (JSON "DB")

1. **වෙනම private repo** එකක් හදන්න — උදා: `hkeflpl-data`
   (app code එකේ repo එකම නෙවෙයි; data වෙනම තියෙන එක පිරිසිදුයි).
2. **Fine-grained PAT** එකක්: GitHub → Settings → Developer settings →
   Personal access tokens → *Fine-grained tokens*.
   - Repository access: ඒ `hkeflpl-data` repo එක විතරක්.
   - Permissions → **Contents: Read and write**.
3. `.streamlit/secrets.toml` එකේ `GITHUB_TOKEN`, `GITHUB_REPO`,
   `JSON_PATH` දාන්න. JSON file එක තවම නැතත් කමක් නෑ — පළවෙනි
   save එකේදී auto හැදෙයි.
4. Sidebar එකේ **Connection test** එබුවම status එක පෙනේවි.

## 3. Streamlit Community Cloud deploy

1. App code එක (data repo එක නෙවෙයි) GitHub එකකට push කරන්න.
2. share.streamlit.io → New app → repo + `app.py` select.
3. **App settings → Secrets** එකට `secrets.toml.example` content එක
   (real values සමග) paste කරන්න. `secrets.toml` file එක push **කරන්න එපා**.

## JSON structure

```json
{
  "schema_version": 1,
  "pl_current": [ { "...23 PL columns..." } ],
  "history": [
    { "run_id", "timestamp", "user", "load_ids", "sl_flag",
      "record_count", "pl_records": [ ... ] }
  ],
  "comparison_summaries": [
    { "run_id", "timestamp", "user", "total", "mismatches",
      "rows": [ { "ID", "Cus Qty", "Inv Qty", "Status" } ] }
  ]
}
```

## සැලකිලිමත් වෙන්න ඕන දේවල් (GitHub-JSON = DB tradeoffs)

- **Concurrency:** 2-3 දෙනෙක්ට හරි. එකවර write දෙකක් ආවොත් optimistic
  locking + retry (5 වතාවක්) වැඩ කරනවා. ගොඩක් එකවර writes නම් 409
  conflicts එයි.
- **Speed:** හැම save එකක්ම git commit එකක් — තප්පර 1-2.
- **Size:** JSON එක ~1 MB යටින් තියාගන්න. `github_store.MAX_HISTORY`
  (default 100) එකෙන් පරණ runs trim වෙනවා.
- **Secrets:** PAT එක **කවදාවත් commit කරන්න එපා**.

## Column names (auto-detect)

**Header row auto-detect:** Step 1 එකේ app එක මුල් rows කිහිපය scan කරලා
real header row එක තෝරගන්නවා — පරණ export එකේ (title row උඩින්) වුණත්,
අලුත් export එකේ (header එක row 1) වුණත් වැඩ. ඕන නම් manual override
(header row number) එකක් තියෙනවා.

`pipeline.find_col` එකෙන් case-insensitive matching කරනවා. නියම WMS
inventory export එකේ headers (verified): `Pallet, Lot Number, Actual Qty,
Color, Size, Style, Supplier, Plant, Item Id, Current Net Weight,
Current Gross Weight, Supplier Desc, Vendor Name, Load Id, Po Number,
Asn Number, Customer Ref Number, Client So` ... (62 columns).

- **`Supplier`** = barcode/EAN (උදා `657001359806`) → PL **SKU BARCODE** +
  comparison key.
- **`Supplier Desc`** = composition text → PL **Composition**.
- **`Plant`** (උදා `9572~AW26 BULK`) → `~` වලින් split → **SO & PO** + **Season**.
- **`Style`** = product name → PL **Article No** (VBA mapping එක; ඕන නම්
  වෙනස් කරන්න `pipeline.generate_pl`).

**Load ID filter column selector:** ඇත්ත export එකේ `Load Id` හිස්
වෙන්න පුළුවන්. Step 3 එකේ **Filter column** dropdown එකෙන් ඇත්තටම
filter කරන column එක (`Order Number` / `Po Number` / `Asn Number` /
`Customer Ref Number` වගේ) තෝරන්න පුළුවන්.

### Supported source formats

App එක source export 2ක්ම handle කරනවා (header auto-detect + alias matching):

1. **Inventory Report** — `Pallet`, `Supplier Desc` (composition),
   `Current Net/Gross Weight`, `Vendor Name` තියෙනවා.
2. **Outbound Staged Licence Plate** — `Pallet` වෙනුවට **`Hu Id`**
   (licence plate) → Ref No.; **`Order Number`** → SO Cust Order Number
   (Cust PO Number වෙනුවට); `Supplier` එකේ GRN suffix (`...-GRN-...`)
   barcode එකට clean වෙනවා; `Supplier Desc`/weights/`Vendor Name` නැති
   නිසා ඒ columns blank වෙනවා.

### Enrichment — blank data වෙනම Inventory එකකින් පුරවන්න

Step 1 එකේ **Enrichment Inventory (optional)** uploader එකක් තියෙනවා.
Main report එකේ (උදා Outbound) `Supplier Desc` / `Current Net&Gross
Weight` / `Vendor Name` හිස් නම්, full Inventory export එකක් upload කරලා
ඒ values match කරලා පුරවගන්න පුළුවන්:

- **Match key** එක දෙ-file එකේම තියෙන column එකක් (default `Supplier`
  = barcode; value overlap වැඩිම එක මුලින් suggest වෙනවා).
- `Supplier` barcode එකේ GRN suffix normalize වෙලා match වෙනවා
  (`657001294220-GRN-1794391` ≡ `657001294220`).
- හිස් cells විතරක් පිරෙනවා; දැනට data තියෙන cells වෙනස් වෙන්නේ නෑ.
- කී rows ක් filled වුණාද කියලා Step 1 එකේ පෙන්නනවා.



