
import streamlit as st
import pandas as pd
import io
import json
from io import BytesIO
from datetime import datetime
import os
from pathlib import Path
import hashlib

st.set_page_config(page_title="BHI PO → Invoices → Items (v8.2)", page_icon="📑", layout="wide")

# ============ Storage paths (local persistence) ============
STORAGE_DIR = Path("storage")
ACTIVE_XLSX = STORAGE_DIR / "active_workbook.xlsx"
ACTIVE_META = STORAGE_DIR / "active_meta.json"
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

def _write_active_workbook(file_bytes: bytes, original_name: str):
    ACTIVE_XLSX.write_bytes(file_bytes)
    meta = {
        "original_name": original_name,
        "saved_at_utc": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
        "size_bytes": len(file_bytes),
        "sha256": hashlib.sha256(file_bytes).hexdigest(),
    }
    ACTIVE_META.write_text(json.dumps(meta, indent=2), encoding="utf-8")

def _read_active_meta():
    if ACTIVE_META.exists():
        try:
            return json.loads(ACTIVE_META.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None

# ============ Optional simple auth ============
_APP_PASSWORD = None
try:
    _APP_PASSWORD = st.secrets.get("APP_PASSWORD")
except Exception:
    _APP_PASSWORD = os.environ.get("APP_PASSWORD")

if _APP_PASSWORD:
    pw = st.text_input("Enter access password", type="password")
    if pw != _APP_PASSWORD:
        st.stop()

# ============ Styles ============
st.markdown(
    """
    <style>
      .kpi-row { position: sticky; top: 0; z-index: 5; background: var(--background-color); padding-top: .5rem; padding-bottom: .5rem; }
      .status-chip { padding: 2px 8px; border-radius: 999px; font-weight: 600; font-size: 0.85rem; display: inline-block; }
      .chip-open { background:#eef2ff; color:#3730a3; }
      .chip-partial { background:#fff7ed; color:#9a3412; }
      .chip-closed { background:#ecfdf5; color:#065f46; }
      .chip-over { background:#fef2f2; color:#991b1b; }
      .meta-box { padding:.5rem .75rem; background: var(--secondary-background-color); border-radius: .5rem; }
      .muted { color: #6b7280; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============ Helpers ============
def _num_series(s: pd.Series) -> pd.Series:
    if s is None or len(s) == 0:
        return pd.Series(dtype="float64")
    s = s.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _hash_bytes(b: bytes) -> str:
    h = hashlib.sha256(); h.update(b); return h.hexdigest()

def get_query_params():
    if hasattr(st, "query_params"):
        return st.query_params
    else:
        return st.experimental_get_query_params()

def set_query_params(**kwargs):
    if hasattr(st, "query_params"):
        st.query_params.clear()
        for k, v in kwargs.items():
            if v is None: continue
            st.query_params[k] = v
    else:
        st.experimental_set_query_params(**{k: v for k, v in kwargs.items() if v is not None})

def paginate(df: pd.DataFrame, key: str, per_page: int = 100) -> pd.DataFrame:
    total = len(df)
    if total <= per_page:
        return df
    pages = max(1, (total - 1) // per_page + 1)
    col1, col2 = st.columns([1,6])
    page = col1.number_input("Page", 1, pages, key=f"{key}_page")
    start = (page - 1) * per_page
    col2.caption(f"Showing {start+1}-{min(start+per_page, total)} of {total}")
    return df.iloc[start:start+per_page]

# Cache compatibility
cache_data = getattr(st, "cache_data", st.cache)

@cache_data
def load_workbook_bytes(b_hash: str, content: bytes):
    xl = pd.ExcelFile(io.BytesIO(content))
    dfs = {s: xl.parse(s, dtype=object) for s in xl.sheet_names}
    return {k: _normalize_cols(v) for k, v in dfs.items()}

@cache_data
def load_workbook_path(path_str: str, mtime: float, size: int):
    # mtime & size included to bust cache when file updates
    xl = pd.ExcelFile(path_str)
    dfs = {s: xl.parse(s, dtype=object) for s in xl.sheet_names}
    return {k: _normalize_cols(v) for k, v in dfs.items()}

def get_dfs(uploaded):
    """
    Priority:
    1) If new upload provided, load it. If 'Persist upload' is checked, save to storage and use going forward.
    2) Else if stored active workbook exists, load it.
    3) Else try local BHI.xlsx beside the app (if present).
    """
    # 1) Uploaded now?
    if uploaded is not None:
        content = uploaded.getvalue()
        if st.session_state.get("persist_upload", True):
            _write_active_workbook(content, uploaded.name)
            st.success(f"📦 Stored as active workbook: {uploaded.name}")
        return load_workbook_bytes(_hash_bytes(content), content)

    # 2) Stored active workbook?
    if ACTIVE_XLSX.exists():
        stat = ACTIVE_XLSX.stat()
        return load_workbook_path(str(ACTIVE_XLSX), stat.st_mtime, stat.st_size)

    # 3) Fallback to BHI.xlsx in repo folder
    if Path("BHI.xlsx").exists():
        stat = Path("BHI.xlsx").stat()
        return load_workbook_path("BHI.xlsx", stat.st_mtime, stat.st_size)

    st.error("No workbook available. Upload a .xlsx or place BHI.xlsx alongside the app.")
    st.stop()

# Polished Excel writer
def build_po_pack_excel(po_num: str, inv_df: pd.DataFrame, items_df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"
    with pd.ExcelWriter(buf, engine=engine) as xw:
        inv_df.to_excel(xw, index=False, sheet_name="Invoices")
        (items_df if items_df is not None and not items_df.empty else pd.DataFrame()).to_excel(
            xw, index=False, sheet_name="Items"
        )
        try:
            ws_inv = xw.sheets["Invoices"]; ws_items = xw.sheets["Items"]
            for ws, df in ((ws_inv, inv_df), (ws_items, (items_df if items_df is not None else pd.DataFrame()))):
                if ws is None: continue
                try: ws.freeze_panes(1,0)
                except Exception: pass
                if df is not None and not df.empty:
                    for i, col in enumerate(df.columns, start=1):
                        try: width = int(df[col].astype(str).str.len().quantile(0.90)) + 2
                        except Exception: width = 12
                        width = max(10, min(48, width))
                        try: ws.set_column(i-1, i-1, width)
                        except Exception:
                            try:
                                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
                            except Exception: pass
        except Exception: pass
    buf.seek(0)
    return buf

def status_with_tol(po_val, inv_val, tol_val):
    po_val = po_val or 0.0
    inv_val = inv_val or 0.0
    if abs(po_val - inv_val) <= tol_val:
        return "CLOSED"
    if inv_val == 0:
        return "OPEN"
    if 0 < inv_val < po_val:
        return "PARTIAL"
    if inv_val > po_val:
        return "OVER-INVOICED"
    return "OPEN"

def status_chip_text(s: str) -> str:
    s = (s or "").upper()
    if s == "CLOSED":  return '<span class="status-chip chip-closed">CLOSED</span>'
    if s == "PARTIAL": return '<span class="status-chip chip-partial">PARTIAL</span>'
    if s == "OVER-INVOICED": return '<span class="status-chip chip-over">OVER</span>'
    return '<span class="status-chip chip-open">OPEN</span>'

# ============ Header ============
st.title("📑 PO → Invoices → Items (v8.2)")
st.caption("Adds **persistent workbook storage** to v8.1 (keeps last uploaded .xlsx until you replace it).")

# Upload & persist controls
meta = _read_active_meta()
with st.expander("📦 Active workbook storage", expanded=True):
    colA, colB, colC = st.columns([3,2,2])
    with colA:
        st.checkbox("Persist upload (make it active)", value=True, key="persist_upload",
                    help="If checked, newly uploaded file is saved to app storage and reused next time.")
        uploaded = st.file_uploader("Upload a .xlsx", type=["xlsx"], accept_multiple_files=False)
    with colB:
        if ACTIVE_XLSX.exists():
            st.button("Clear stored workbook", type="secondary", key="clear_store")
            if st.session_state.get("clear_store"):
                try:
                    ACTIVE_XLSX.unlink(missing_ok=True)
                    ACTIVE_META.unlink(missing_ok=True)
                    st.success("Cleared stored workbook.")
                except Exception as e:
                    st.error(f"Could not clear: {e}")
                # reset flag so it doesn't retrigger
                st.session_state["clear_store"] = False
    with colC:
        if meta and ACTIVE_XLSX.exists():
            st.markdown(
                f"""
                <div class="meta-box">
                  <div><b>Stored file:</b> {meta.get("original_name","active_workbook.xlsx")}</div>
                  <div class="muted">Saved (UTC): {meta.get("saved_at_utc","")}</div>
                  <div class="muted">Size: {meta.get("size_bytes",0):,} bytes</div>
                  <div class="muted">SHA256: {meta.get("sha256","")[:12]}…</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        elif ACTIVE_XLSX.exists():
            st.info("An active workbook is stored (metadata unavailable).")

# Load workbook dict (sheet_name → DataFrame)
dfs = get_dfs(uploaded)

# Identify sheets (with fuzzy matches)
def find_sheet_like(d: dict, name: str):
    if name in d: return name
    for k in d.keys():
        if name.lower() in k.lower().replace(" ", ""):
            return k
    return None

sheet_pos = find_sheet_like(dfs, "POs") or (list(dfs.keys())[0] if dfs else None)
sheet_inv = find_sheet_like(dfs, "Invoices") or (list(dfs.keys())[1] if len(dfs) > 1 else sheet_pos)
sheet_items = find_sheet_like(dfs, "InvoiceItems")
sheet_poitems = find_sheet_like(dfs, "POItems") or find_sheet_like(dfs, "PO_Items") or find_sheet_like(dfs, "POLines")

POs = dfs.get(sheet_pos, pd.DataFrame()).copy() if sheet_pos else pd.DataFrame()
Invoices = dfs.get(sheet_inv, pd.DataFrame()).copy() if sheet_inv else pd.DataFrame()
InvoiceItems = dfs.get(sheet_items, pd.DataFrame()) if sheet_items else pd.DataFrame()
POItems = dfs.get(sheet_poitems, pd.DataFrame()) if sheet_poitems else pd.DataFrame()

for df in (POs, Invoices, InvoiceItems, POItems):
    if not df.empty:
        for col in df.columns:
            if pd.api.types.is_object_dtype(df[col]):
                df[col] = df[col].astype(str).str.strip()

# Sidebar: basics
st.sidebar.header("Settings")
tol = st.sidebar.number_input("Close tolerance (amount)", value=0.01, step=0.01)

# Required columns (soft check)
def warn_required(name, df, cols):
    need = [c for c in cols if c not in df.columns]
    if need: st.warning(f"'{name}' missing likely columns: {', '.join(need)}")
warn_required("POs", POs, ["PO_NUMBER", "PO_AMOUNT"])
warn_required("Invoices", Invoices, ["INVOICE_NUMBER", "PO_NUMBER", "INVOICE_AMOUNT"])

# PO search & deep link
qp = get_query_params()
default_po = None
if isinstance(qp, dict) and "po" in qp:
    default_po = qp["po"] if isinstance(qp["po"], str) else (qp["po"][0] if qp["po"] else None)

if "PO_NUMBER" not in POs.columns or POs.empty:
    st.error("POs sheet is missing or has no 'PO_NUMBER'."); st.stop()

left, right = st.columns([2,3])
with left:
    po_query = st.text_input("Find PO (partial or full)", value=default_po or "", placeholder="e.g., 9460").strip()

po_list = POs["PO_NUMBER"].dropna().astype(str).unique().tolist()
matches = [p for p in po_list if po_query.lower() in p.lower()] if po_query else po_list
matches = matches[:500]
if not matches:
    st.warning("No POs match that search. Try a different query.")
    st.stop()
default_index = 0 if default_po and default_po in matches else 0
picked_po = st.selectbox("Select PO", options=matches, index=default_index)
set_query_params(po=picked_po)

if not picked_po:
    st.info("Type to search or pick a PO."); st.stop()

# Invoices for PO
if "PO_NUMBER" not in Invoices.columns:
    st.error("Invoices sheet missing 'PO_NUMBER'."); st.stop()
inv_for_po = Invoices[Invoices["PO_NUMBER"].astype(str) == str(picked_po)].copy()

amt_ser = _num_series(inv_for_po["INVOICE_AMOUNT"]) if "INVOICE_AMOUNT" in inv_for_po.columns else pd.Series(dtype="float64")
total_invoiced = float(amt_ser.sum()) if not amt_ser.empty else 0.0

po_amount = None
if {"PO_NUMBER","PO_AMOUNT"}.issubset(POs.columns):
    r = POs[POs["PO_NUMBER"].astype(str) == str(picked_po)]
    if not r.empty:
        po_amount = float(_num_series(pd.Series([r.iloc[0]["PO_AMOUNT"]])).fillna(0).iloc[0])

st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
k1, k2, k3 = st.columns(3)
k1.metric("Invoiced", f"{total_invoiced:,.0f}")
if po_amount is not None:
    k2.metric("PO Amount", f"{po_amount:,.0f}")
    k3.metric("Variance", f"{(po_amount - total_invoiced):,.0f}")
st.markdown('</div>', unsafe_allow_html=True)

# ============ Tabs ============
tabs = ["Invoices", "Items (from invoices)"]
if not POItems.empty:
    tabs.append("Uninvoiced PO Items")
tab_objs = st.tabs(tabs)

# Invoices tab
with tab_objs[0]:
    st.subheader("Invoices for this PO")
    inv_quick = st.text_input("Filter invoices (contains, any column)", key="inv_filter").lower().strip()
    if inv_quick:
        inv_for_po = inv_for_po[inv_for_po.apply(lambda r: r.astype(str).str.lower().str.contains(inv_quick, na=False)).any(axis=1)]
    default_inv_cols = [c for c in ["INVOICE_NUMBER","INVOICE_DATE","INVOICE_AMOUNT","CURRENCY","STATUS","PO_NUMBER"] if c in inv_for_po.columns]
    pick_cols_inv = st.multiselect("Columns to show", list(inv_for_po.columns), default=default_inv_cols, key="inv_cols_show")
    view_df = inv_for_po[pick_cols_inv] if pick_cols_inv else inv_for_po
    st.dataframe(paginate(view_df, key="inv", per_page=100), use_container_width=True)

    # Excel pack (invoices + their items)
    inv_items_for_po = InvoiceItems.merge(Invoices[["INVOICE_NUMBER","PO_NUMBER"]], on="INVOICE_NUMBER", how="left") if not InvoiceItems.empty else pd.DataFrame()
    po_items_full = inv_items_for_po[inv_items_for_po["PO_NUMBER"].astype(str) == str(picked_po)] if not inv_items_for_po.empty else pd.DataFrame()
    xbuf = build_po_pack_excel(str(picked_po), inv_for_po, po_items_full)
    st.download_button("⬇️ Download PO Pack (Excel)", data=xbuf, file_name=f"{picked_po}_pack.xlsx")

# Items-from-invoices tab
with tab_objs[1]:
    st.subheader("Items (from invoices)")
    inv_items_for_po = InvoiceItems.merge(Invoices[["INVOICE_NUMBER","PO_NUMBER"]], on="INVOICE_NUMBER", how="left") if not InvoiceItems.empty else pd.DataFrame()
    po_items_full = inv_items_for_po[inv_items_for_po["PO_NUMBER"].astype(str) == str(picked_po)] if not inv_items_for_po.empty else pd.DataFrame()
    items_pref = ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"]
    show_cols = [c for c in items_pref if c in po_items_full.columns]
    if po_items_full.empty:
        st.info("No invoice items found for this PO.")
    else:
        st.dataframe(paginate(po_items_full[show_cols] if show_cols else po_items_full, key="items", per_page=200), use_container_width=True)
        st.download_button("⬇️ Download PO Items (CSV)",
                           data=(po_items_full[show_cols] if show_cols else po_items_full).to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{picked_po}_all_items.csv", mime="text/csv")

# Uninvoiced PO items tab (if POItems provided)
if not POItems.empty and len(tab_objs) >= 3:
    with tab_objs[2]:
        st.subheader("Uninvoiced PO Items")
        # Column matching
        def get_col(df, candidates):
            for c in candidates:
                if c in df.columns:
                    return c
            return None

        poi_po = get_col(POItems, ["PO_NUMBER","PO No","PONUMBER","PO"])
        poi_line = get_col(POItems, ["LINE","LINE_NO","ITEM","ITEM_NO","PO_LINE"])
        poi_qty = get_col(POItems, ["QTY","ORDER_QTY","QUANTITY"])
        poi_price = get_col(POItems, ["UNIT_PRICE","PRICE","UNITPRICE"])
        poi_total = get_col(POItems, ["LINE_TOTAL","TOTAL","AMOUNT","LINEAMOUNT"])
        poi_mat = get_col(POItems, ["MATERIAL","SKU","ITEM_CODE"])
        poi_desc = get_col(POItems, ["DESCRIPTION","DESC"])

        if poi_po is None or poi_line is None:
            st.error("POItems must include PO number and line columns (e.g., PO_NUMBER and LINE).")
        else:
            po_lines = POItems[POItems[poi_po].astype(str) == str(picked_po)].copy()
            if po_lines.empty:
                st.info("No PO lines found for this PO in POItems sheet.")
            else:
                if poi_total is None and (poi_qty is not None and poi_price is not None):
                    po_lines["_PO_TOTAL"] = _num_series(po_lines[poi_qty]) * _num_series(po_lines[poi_price])
                else:
                    po_lines["_PO_TOTAL"] = _num_series(po_lines[poi_total]) if poi_total in po_lines.columns else pd.Series([None]*len(po_lines))

                # Map invoice items → PO
                inv_items = pd.DataFrame()
                if not InvoiceItems.empty and {"INVOICE_NUMBER","PO_NUMBER"}.issubset(Invoices.columns):
                    inv_items = InvoiceItems.merge(Invoices[["INVOICE_NUMBER","PO_NUMBER"]], on="INVOICE_NUMBER", how="left")
                    inv_items = inv_items[inv_items["PO_NUMBER"].astype(str) == str(picked_po)]

                ii_line = get_col(inv_items, ["LINE","LINE_NO","ITEM","ITEM_NO","PO_LINE"])
                ii_qty = get_col(inv_items, ["QTY","QUANTITY"])
                ii_total = get_col(inv_items, ["LINE_TOTAL","TOTAL","AMOUNT"])
                ii_price = get_col(inv_items, ["UNIT_PRICE","PRICE"])

                if not inv_items.empty and ii_line is not None:
                    billed = inv_items.copy()
                    billed["_BILLED_QTY"] = _num_series(billed[ii_qty]) if ii_qty in billed.columns else 0
                    if ii_total in billed.columns:
                        billed["_BILLED_AMT"] = _num_series(billed[ii_total])
                    else:
                        billed["_BILLED_AMT"] = _num_series(billed[ii_qty]) * _num_series(billed[ii_price]) if (ii_qty in billed.columns and ii_price in billed.columns) else 0
                    agg = billed.groupby(ii_line, dropna=False).agg({"_BILLED_QTY":"sum","_BILLED_AMT":"sum"}).reset_index().rename(columns={ii_line:"__LINE_KEY"})
                else:
                    agg = pd.DataFrame({"__LINE_KEY": po_lines[poi_line].astype(str).unique().tolist(), "_BILLED_QTY":[0]*po_lines[poi_line].nunique(), "_BILLED_AMT":[0]*po_lines[poi_line].nunique()})

                po_lines["__LINE_KEY"] = po_lines[poi_line].astype(str)
                merged = po_lines.merge(agg, on="__LINE_KEY", how="left")
                merged["_BILLED_QTY"] = merged["_BILLED_QTY"].fillna(0)
                merged["_BILLED_AMT"] = merged["_BILLED_AMT"].fillna(0)

                if poi_qty in merged.columns:
                    merged["_PO_QTY"] = _num_series(merged[poi_qty]).fillna(0)
                    merged["REMAIN_QTY"] = (merged["_PO_QTY"] - merged["_BILLED_QTY"]).clip(lower=0)
                else:
                    merged["REMAIN_QTY"] = None
                merged["REMAIN_AMT"] = (merged["_PO_TOTAL"].fillna(0) - merged["_BILLED_AMT"]).clip(lower=0)

                remain = merged[[c for c in [poi_line, poi_mat, poi_desc, poi_qty, poi_price, poi_total, "_PO_TOTAL", "_BILLED_QTY", "_BILLED_AMT", "REMAIN_QTY", "REMAIN_AMT"] if c in merged.columns]].copy()
                remain = remain[(remain["REMAIN_QTY"].fillna(0) > 0) | (remain["REMAIN_AMT"].fillna(0) > 0)]

                if remain.empty:
                    st.success("All PO lines appear fully invoiced for this PO 🎉")
                else:
                    remain = remain.rename(columns={"_PO_TOTAL":"PO_LINE_TOTAL","_BILLED_QTY":"BILLED_QTY","_BILLED_AMT":"BILLED_AMT"})
                    st.dataframe(paginate(remain, key="uninvoiced", per_page=200), use_container_width=True)
                    st.download_button("⬇️ Download Uninvoiced PO Items (CSV)",
                                       data=remain.to_csv(index=False).encode("utf-8-sig"),
                                       file_name=f"{picked_po}_uninvoiced_po_items.csv", mime="text/csv")

# Footer
st.caption(f"Built {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')} • v8.2 (persistent storage)")
