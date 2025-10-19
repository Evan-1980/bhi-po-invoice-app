
import streamlit as st
import pandas as pd
import io
import hashlib
from io import BytesIO
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="BHI PO → Invoices → Items (v8)", page_icon="📑", layout="wide")

# ===================== Optional simple auth (safe) =====================
_APP_PASSWORD = None
try:
    _APP_PASSWORD = st.secrets.get("APP_PASSWORD")
except Exception:
    _APP_PASSWORD = os.environ.get("APP_PASSWORD")

if _APP_PASSWORD:
    pw = st.text_input("Enter access password", type="password")
    if pw != _APP_PASSWORD:
        st.stop()

# ===================== Sticky KPI + chip styles =====================
st.markdown(
    """
    <style>
      .kpi-row { position: sticky; top: 0; z-index: 5; background: var(--background-color); padding-top: .5rem; padding-bottom: .5rem; }
      .status-chip { padding: 2px 8px; border-radius: 999px; font-weight: 600; font-size: 0.85rem; display: inline-block; }
      .chip-open { background:#eef2ff; color:#3730a3; }
      .chip-partial { background:#fff7ed; color:#9a3412; }
      .chip-closed { background:#ecfdf5; color:#065f46; }
      .chip-over { background:#fef2f2; color:#991b1b; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ===================== Helpers =====================
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
def load_workbook_local(path: str):
    xl = pd.ExcelFile(path)
    dfs = {s: xl.parse(s, dtype=object) for s in xl.sheet_names}
    return {k: _normalize_cols(v) for k, v in dfs.items()}

@cache_data
def load_workbook_bytes(b_hash: str, content: bytes):
    xl = pd.ExcelFile(io.BytesIO(content))
    dfs = {s: xl.parse(s, dtype=object) for s in xl.sheet_names}
    return {k: _normalize_cols(v) for k, v in dfs.items()}

def get_dfs(uploaded):
    if uploaded is not None:
        content = uploaded.getvalue()
        return load_workbook_bytes(_hash_bytes(content), content)
    try:
        return load_workbook_local("BHI.xlsx")
    except Exception:
        st.error("No file uploaded and `BHI.xlsx` not found alongside the app.")
        st.stop()

# Polished Excel writer with autosizing and freeze header; engine fallback
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
            ws_inv = xw.sheets["Invoices"]
            ws_items = xw.sheets["Items"]
            for ws, df in ( (ws_inv, inv_df), (ws_items, (items_df if items_df is not None else pd.DataFrame())) ):
                if ws is None: continue
                try: ws.freeze_panes(1, 0)
                except Exception: pass
                if df is not None and not df.empty:
                    for i, col in enumerate(df.columns, start=1):
                        try:
                            width = int(df[col].astype(str).str.len().quantile(0.90)) + 2
                        except Exception:
                            width = 12
                        width = max(10, min(48, width))
                        try:
                            ws.set_column(i-1, i-1, width)
                        except Exception:
                            try:
                                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
                            except Exception:
                                pass
        except Exception:
            pass
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

# ===================== UI: File =====================
st.title("📑 PO → Invoices → Items (v8)")
st.caption("Search PO, view invoices, items, exports, quality checks, suggestions, and management rollups.")

uploaded = st.file_uploader("Upload a .xlsx (optional). If omitted, the app will load BHI.xlsx from the app folder.", type=["xlsx"])
dfs = get_dfs(uploaded)

# Identify sheets
sheet_pos = "POs" if "POs" in dfs else (list(dfs.keys())[0] if dfs else None)
sheet_inv = "Invoices" if "Invoices" in dfs else (list(dfs.keys())[1] if len(dfs) > 1 else sheet_pos)
sheet_items = "InvoiceItems" if "InvoiceItems" in dfs else None

POs = dfs.get(sheet_pos, pd.DataFrame()).copy() if sheet_pos else pd.DataFrame()
Invoices = dfs.get(sheet_inv, pd.DataFrame()).copy() if sheet_inv else pd.DataFrame()
InvoiceItems = dfs.get(sheet_items, pd.DataFrame()) if sheet_items else pd.DataFrame()

for df in (POs, Invoices, InvoiceItems):
    if not df.empty:
        for col in df.columns:
            if pd.api.types.is_object_dtype(df[col]):
                df[col] = df[col].astype(str).str.strip()

# Sidebar: settings
st.sidebar.header("Settings")
tol = st.sidebar.number_input("Close tolerance (amount)", value=0.01, step=0.01)
st.sidebar.subheader("Currency normalization")
base_ccy = st.sidebar.text_input("Base currency", value="USD")
rates_raw = st.sidebar.text_area("Rates (one per line, e.g. 'IQD=0.00076')", value="IQD=0.00076\nEUR=1.09")
rates = {}
for line in rates_raw.splitlines():
    if "=" in line:
        k, v = line.split("=", 1)
        try:
            rates[k.strip().upper()] = float(v)
        except Exception:
            pass
def normalize_amount(amount_series: pd.Series, ccy_series: pd.Series) -> pd.Series:
    a = _num_series(amount_series)
    fx = ccy_series.astype(str).str.upper().map(rates).fillna(1.0)
    return a * fx
show_quality = st.sidebar.checkbox("Show Data Quality tab", value=True)

# Required columns check (soft)
REQUIRED = {
    "POs": ["PO_NUMBER", "PO_AMOUNT"],
    "Invoices": ["INVOICE_NUMBER", "PO_NUMBER", "INVOICE_AMOUNT"],
    "InvoiceItems": ["INVOICE_NUMBER"]
}
def warn_required(name, df):
    need = [c for c in REQUIRED[name] if c not in df.columns]
    if need:
        st.warning(f"'{name}' missing likely columns: {', '.join(need)}")
for name, df in (("POs", POs), ("Invoices", Invoices)):
    warn_required(name, df)
if not InvoiceItems.empty:
    warn_required("InvoiceItems", InvoiceItems)

# PO search & deep link
qp = get_query_params()
default_po = None
if isinstance(qp, dict) and "po" in qp:
    default_po = qp["po"] if isinstance(qp["po"], str) else (qp["po"][0] if qp["po"] else None)

left, right = st.columns([2, 3])
with left:
    po_query = st.text_input("Find PO (partial or full)", value=default_po or "", placeholder="e.g., 9460").strip()

if "PO_NUMBER" not in POs.columns or POs.empty:
    st.error("POs sheet is missing or has no 'PO_NUMBER'."); st.stop()

po_list = POs["PO_NUMBER"].dropna().astype(str).unique().tolist()
matches = [p for p in po_list if po_query.lower() in p.lower()] if po_query else po_list
matches = matches[:500]
default_index = 0 if default_po and default_po in matches else (0 if matches else None)
picked_po = st.selectbox("Select PO", options=matches, index=default_index)
set_query_params(po=picked_po)

if not picked_po:
    st.info("Type to search or pick a PO.")
    st.stop()

# Invoices for selected PO
if "PO_NUMBER" not in Invoices.columns:
    st.error("Invoices sheet missing 'PO_NUMBER'."); st.stop()
inv_for_po = Invoices[Invoices["PO_NUMBER"].astype(str) == str(picked_po)].copy()

# KPIs (sticky)
amt_ser = _num_series(inv_for_po["INVOICE_AMOUNT"]) if "INVOICE_AMOUNT" in inv_for_po.columns else pd.Series(dtype="float64")
total_invoiced = float(amt_ser.sum()) if not amt_ser.empty else 0.0

total_invoiced_norm = None
if {"INVOICE_AMOUNT","CURRENCY"}.issubset(inv_for_po.columns):
    total_invoiced_norm = float(normalize_amount(inv_for_po["INVOICE_AMOUNT"], inv_for_po["CURRENCY"]).sum())

po_amount = None
if {"PO_NUMBER","PO_AMOUNT"}.issubset(POs.columns):
    r = POs[POs["PO_NUMBER"].astype(str) == str(picked_po)]
    if not r.empty:
        po_amount = float(_num_series(pd.Series([r.iloc[0]["PO_AMOUNT"]])).fillna(0).iloc[0])

st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
k1, k2, k3, k4 = st.columns(4)
k1.metric("Invoiced", f"{total_invoiced:,.0f}")
if total_invoiced_norm is not None:
    k2.metric(f"Invoiced → {base_ccy}", f"{total_invoiced_norm:,.0f}")
if po_amount is not None:
    k3.metric("PO Amount", f"{po_amount:,.0f}")
    k4.metric("Variance", f"{(po_amount - total_invoiced):,.0f}")
st.markdown('</div>', unsafe_allow_html=True)

if po_amount is not None:
    st.caption(f"Status (tol={tol}): {status_chip_text(status_with_tol(po_amount, total_invoiced, tol))}", unsafe_allow_html=True)

# Tabs
tabs = ["Invoices", "Items", "Rollups", "Suggest", "Quality"] if show_quality else ["Invoices", "Items", "Rollups", "Suggest"]
tab_inv, tab_items, tab_rollups, tab_suggest, *rest = st.tabs(tabs)
tab_quality = rest[0] if rest else None

# ===== Invoices Tab =====
with tab_inv:
    st.subheader("Invoices for this PO")
    inv_quick = st.text_input("Filter invoices (contains, any column)", key="inv_filter").lower().strip()
    if inv_quick:
        inv_for_po = inv_for_po[inv_for_po.apply(lambda r: r.astype(str).str.lower().str.contains(inv_quick, na=False)).any(axis=1)]

    # Colored status chip column
    if po_amount is not None:
        # Compute per-invoice status vs PO share (coarse): not exact allocation, but gives visibility
        inv_for_po["_amt"] = _num_series(inv_for_po.get("INVOICE_AMOUNT", pd.Series(dtype=str)))
        inv_for_po["_chip"] = inv_for_po["_amt"].apply(lambda x: status_chip_text("PARTIAL" if 0 < (x or 0) < (po_amount or 0) else ("OVER-INVOICED" if (x or 0) > (po_amount or 0) else ("CLOSED" if abs((po_amount or 0)-(x or 0))<=tol else "OPEN"))))
    else:
        inv_for_po["_chip"] = status_chip_text("OPEN")

    inv_cols_available = [c for c in inv_for_po.columns if not c.startswith("_")]
    default_inv_cols = [c for c in ["INVOICE_NUMBER","INVOICE_DATE","INVOICE_AMOUNT","CURRENCY","STATUS","PO_NUMBER"] if c in inv_cols_available]
    pick_cols_inv = st.multiselect("Columns to show", inv_cols_available, default=default_inv_cols, key="inv_cols_show")

    # Show with chips
    show_df = inv_for_po[[c for c in pick_cols_inv if c in inv_for_po.columns]].copy() if pick_cols_inv else inv_for_po.drop(columns=[c for c in inv_for_po.columns if c.startswith("_")], errors="ignore").copy()
    show_df.insert(0, "STATUS_CHIP", inv_for_po["_chip"])
    paged = paginate(show_df, key="inv", per_page=100)
    st.write(paged.to_html(escape=False, index=False), unsafe_allow_html=True)

    # Invoice totals summary
    if "INVOICE_NUMBER" in inv_for_po.columns:
        summary = (inv_for_po.assign(_amt=_num_series(inv_for_po.get("INVOICE_AMOUNT", pd.Series(dtype=str))))
                   .groupby("INVOICE_NUMBER", dropna=True)["_amt"].sum().reset_index()
                   .rename(columns={"_amt": "Invoice Total"}))
        st.markdown("**Invoice totals**")
        st.dataframe(summary, use_container_width=True)

    # Exports
    inv_nums = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist() if "INVOICE_NUMBER" in inv_for_po.columns else []
    po_items_full = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str).isin(inv_nums)].copy() if (not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns) else pd.DataFrame()
    xbuf = build_po_pack_excel(str(picked_po), inv_for_po, po_items_full)
    st.download_button("⬇️ Download PO Pack (Excel)", data=xbuf, file_name=f"{picked_po}_pack.xlsx")
    st.download_button("⬇️ Export current invoices (JSON)", data=inv_for_po.to_json(orient="records").encode("utf-8"), file_name=f"invoices_{picked_po}.json", mime="application/json")

# ===== Items Tab =====
with tab_items:
    st.subheader("Items")
    inv_nums = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist() if "INVOICE_NUMBER" in inv_for_po.columns else []
    po_items_full = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str).isin(inv_nums)].copy() if (not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns) else pd.DataFrame()
    show_all_items = st.checkbox("Show all items for this PO", value=True)
    if show_all_items:
        items_pref = ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"]
        show_cols = [c for c in items_pref if c in po_items_full.columns]
        if po_items_full.empty:
            st.info("No items found across invoices for this PO.")
        else:
            paged_items = paginate(po_items_full[show_cols] if show_cols else po_items_full, key="items", per_page=200)
            st.dataframe(paged_items, use_container_width=True)
            st.download_button(
                "⬇️ Download PO Items (CSV)",
                data=(po_items_full[show_cols] if show_cols else po_items_full).to_csv(index=False).encode("utf-8-sig"),
                file_name=f"{picked_po}_all_items.csv",
                mime="text/csv"
            )
    else:
        invoice_numbers = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist() if "INVOICE_NUMBER" in inv_for_po.columns else []
        picked_invoice = st.selectbox("Invoice Number", options=invoice_numbers, index=0 if invoice_numbers else None)
        if picked_invoice and not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns:
            items = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str) == str(picked_invoice)].copy()
            items_pref = ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"]
            show_cols = [c for c in items_pref if c in items.columns]
            if items.empty:
                st.info("No items for the selected invoice.")
            else:
                paged_items = paginate(items[show_cols] if show_cols else items, key="items_one", per_page=200)
                st.dataframe(paged_items, use_container_width=True)
                st.download_button(
                    "⬇️ Download items (CSV)",
                    data=(items[show_cols] if show_cols else items).to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"{picked_invoice}_items.csv",
                    mime="text/csv"
                )

# ===== Rollups Tab =====
with tab_rollups:
    st.subheader("Management rollups")

    # Optional DuckDB acceleration
    duck_ok = False
    try:
        import duckdb as dd  # type: ignore
        duck_ok = True
    except Exception:
        duck_ok = False

    # Vendor summary
    if "VENDOR" in inv_for_po.columns:
        if duck_ok:
            tmp = inv_for_po.assign(_amt=_num_series(inv_for_po.get("INVOICE_AMOUNT", pd.Series(dtype=str))))
            vend = dd.query("SELECT VENDOR, SUM(_amt) AS TotalAmount FROM tmp GROUP BY 1 ORDER BY 2 DESC").to_df()
        else:
            vend = (inv_for_po.assign(_amt=_num_series(inv_for_po.get("INVOICE_AMOUNT", pd.Series(dtype=str))))
                    .groupby("VENDOR", dropna=True)["_amt"].sum().reset_index()
                    .rename(columns={"_amt": "TotalAmount"}).sort_values("TotalAmount", ascending=False))
        st.markdown("**Vendor summary**")
        st.dataframe(vend, use_container_width=True)
    else:
        st.info("No VENDOR column in Invoices for vendor rollup.")

    # Monthly summary
    if "INVOICE_DATE" in inv_for_po.columns:
        tmp = inv_for_po.copy()
        tmp["_d"] = pd.to_datetime(tmp["INVOICE_DATE"], errors="coerce")
        tmp["_amt"] = _num_series(tmp.get("INVOICE_AMOUNT", pd.Series(dtype=str)))
        if duck_ok:
            monthly = dd.query("SELECT strftime(_d, '%Y-%m') AS Month, SUM(_amt) AS Amount FROM tmp GROUP BY 1 ORDER BY 1").to_df()
        else:
            tmp["Month"] = tmp["_d"].dt.to_period("M").astype(str)
            monthly = tmp.groupby("Month")["_amt"].sum().reset_index().rename(columns={"_amt":"Amount"})
        st.markdown("**Monthly invoiced**")
        if not monthly.empty:
            st.bar_chart(monthly.set_index("Month")["Amount"])
            st.dataframe(monthly, use_container_width=True)
        else:
            st.info("No valid invoice dates to summarize.")
    else:
        st.info("No INVOICE_DATE column for monthly summary.")

    # PO status report (all POs)
    def po_status_row(po_num):
        inv = Invoices[Invoices["PO_NUMBER"].astype(str) == str(po_num)]
        total = _num_series(inv.get("INVOICE_AMOUNT", pd.Series(dtype=str))).sum()
        po_amt = None
        if {"PO_NUMBER","PO_AMOUNT"}.issubset(POs.columns):
            row = POs[POs["PO_NUMBER"].astype(str) == str(po_num)]
            if not row.empty:
                po_amt = float(_num_series(pd.Series([row.iloc[0]["PO_AMOUNT"]])).fillna(0).iloc[0])
        stat = status_with_tol(po_amt, total, tol) if po_amt is not None else "UNKNOWN"
        return {"PO_NUMBER": po_num, "PO_AMOUNT": po_amt, "INVOICED": float(total),
                "VARIANCE": None if po_amt is None else float(po_amt - total), "STATUS": stat}

    if st.button("Build PO status report"):
        report = pd.DataFrame([po_status_row(p) for p in po_list])
        st.dataframe(report, use_container_width=True)
        st.download_button("⬇️ Download PO status (CSV)",
                           data=report.to_csv(index=False).encode("utf-8-sig"),
                           file_name="po_status.csv", mime="text/csv")

# ===== Suggest Tab (PO suggestions + 3-way variance) =====
with tab_suggest:
    st.subheader("Suggestions & Controls")

    # 3-way check: Invoice vs Items
    if {"INVOICE_NUMBER"}.issubset(Invoices.columns) and not InvoiceItems.empty:
        items_sum = (InvoiceItems.assign(_line=_num_series(InvoiceItems.get("LINE_TOTAL", pd.Series(dtype=str)))) \
                     .groupby("INVOICE_NUMBER")["_line"].sum().rename("ITEMS_TOTAL"))
        inv_sum   = _num_series(Invoices.get("INVOICE_AMOUNT", pd.Series(dtype=str)))
        chk = Invoices.assign(INV_TOTAL=inv_sum).join(items_sum, on="INVOICE_NUMBER")
        chk["INV_vs_ITEMS"] = chk["INV_TOTAL"] - chk["ITEMS_TOTAL"]
        st.markdown("### 3-way check (Invoice vs Items)")
        st.dataframe(chk[["INVOICE_NUMBER","INV_TOTAL","ITEMS_TOTAL","INV_vs_ITEMS"]].sort_values("INV_vs_ITEMS", key=lambda s: s.abs()), use_container_width=True)
        st.download_button("⬇️ Download 3-way check (CSV)",
                           data=chk.to_csv(index=False).encode("utf-8-sig"),
                           file_name="three_way_check.csv", mime="text/csv")
    else:
        st.info("Need Invoices + InvoiceItems (with LINE_TOTAL) for 3-way check.")

    # PO suggestions for invoices missing PO
    def suggest_po_for_invoice(inv_row, POs, days=7, pct=0.02):
        amt = _num_series(pd.Series([inv_row.get("INVOICE_AMOUNT")])).iloc[0]
        vdr = str(inv_row.get("VENDOR", "")).strip().lower()
        d   = pd.to_datetime(inv_row.get("INVOICE_DATE"), errors="coerce")
        pool = POs.copy()
        if "VENDOR" in pool.columns and vdr:
            pool = pool[pool["VENDOR"].astype(str).str.lower().str.strip() == vdr]
        if "PO_DATE" in pool.columns and pd.notna(d):
            pool["_po_d"] = pd.to_datetime(pool["PO_DATE"], errors="coerce")
            pool = pool[pool["_po_d"].between(d - pd.Timedelta(days, "D"), d + pd.Timedelta(days, "D"))]
        if "PO_AMOUNT" in pool.columns and pd.notna(amt):
            pa = _num_series(pool["PO_AMOUNT"])
            pool = pool[(pa.between(amt*(1-pct), amt*(1+pct)))]
        return pool.head(3)

    if "PO_NUMBER" in Invoices.columns:
        need_link = Invoices[Invoices["PO_NUMBER"].isna()].copy()
        if not need_link.empty:
            st.markdown("### Suggested POs for invoices missing PO")
            sample = need_link.head(20).copy()
            def top_candidates(row):
                cands = suggest_po_for_invoice(row, POs)
                return ",".join(cands["PO_NUMBER"].astype(str)) if not cands.empty else ""
            sample["SUGGESTED_POs"] = sample.apply(top_candidates, axis=1)
            cols = [c for c in ["INVOICE_NUMBER","VENDOR","INVOICE_DATE","INVOICE_AMOUNT","SUGGESTED_POs"] if c in sample.columns or c=="SUGGESTED_POs"]
            st.dataframe(sample[cols], use_container_width=True)
            st.download_button("⬇️ Download suggestions (CSV)",
                               data=sample[cols].to_csv(index=False).encode("utf-8-sig"),
                               file_name="po_suggestions.csv", mime="text/csv")
        else:
            st.success("No invoices missing PO_NUMBER 🎉")

# ===== Quality Tab =====
if show_quality and tab_quality is not None:
    with tab_quality:
        st.subheader("Data Quality")
        issues = []
        actions = []

        dups = pd.DataFrame()
        if "INVOICE_NUMBER" in Invoices.columns:
            vc = Invoices["INVOICE_NUMBER"].astype(str).value_counts(dropna=False)
            dup_keys = vc[vc > 1].index.tolist()
            if dup_keys:
                issues.append(f"Duplicate INVOICE_NUMBERs: {len(dup_keys)} unique values")
                dups = Invoices[Invoices["INVOICE_NUMBER"].astype(str).isin(dup_keys)].copy()
                actions.append(("Duplicates.csv", dups))

        miss_po = pd.DataFrame()
        if "PO_NUMBER" in Invoices.columns:
            miss_po = Invoices[Invoices["PO_NUMBER"].isna()].copy()
            if not miss_po.empty:
                issues.append(f"Invoices missing PO_NUMBER: {len(miss_po)}")
                actions.append(("Invoices_missing_PO.csv", miss_po))

        orphan = pd.DataFrame()
        if "PO_NUMBER" in Invoices.columns and "PO_NUMBER" in POs.columns:
            known_pos = set(POs["PO_NUMBER"].dropna().astype(str))
            orphan = Invoices[~Invoices["PO_NUMBER"].dropna().astype(str).isin(known_pos)].copy()
            if not orphan.empty:
                issues.append(f"Invoices referencing unknown PO_NUMBER: {len(orphan)}")
                actions.append(("Invoices_orphan.csv", orphan))

        miss_inv = pd.DataFrame()
        if not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns:
            miss_inv = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].isna()].copy()
            if not miss_inv.empty:
                issues.append(f"Items missing INVOICE_NUMBER: {len(miss_inv)}")
                actions.append(("Items_missing_invoice.csv", miss_inv))

        zero_neg = pd.DataFrame()
        if "INVOICE_AMOUNT" in Invoices.columns:
            zero_neg = Invoices[_num_series(Invoices["INVOICE_AMOUNT"]).fillna(0) <= 0]
            if not zero_neg.empty:
                issues.append(f"Invoices with zero/negative amount: {len(zero_neg)}")
                actions.append(("Invoices_zero_or_negative.csv", zero_neg))

        if issues:
            for m in issues: st.warning("• " + m)
        else:
            st.success("No obvious issues found.")

        for name, df in actions:
            st.download_button(
                f"⬇️ Download {name}",
                data=df.to_csv(index=False).encode("utf-8-sig"),
                file_name=name,
                mime="text/csv"
            )

# ===== Audit log (simple) =====
if "audit_log" not in st.session_state:
    st.session_state["audit_log"] = []

def log(event: str, detail: str = ""):
    st.session_state["audit_log"].append({
        "ts": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
        "event": event,
        "detail": detail,
        "po": picked_po
    })

# Log key actions
log("open_po", f"Opened PO {picked_po}")
if st.button("⬇️ Download audit log (CSV)"):
    audit_df = pd.DataFrame(st.session_state["audit_log"])
    st.download_button("Download audit.csv", data=audit_df.to_csv(index=False).encode("utf-8-sig"),
                       file_name="audit.csv", mime="text/csv", key="audit_dl_btn")

# Footer
st.caption(f"Built {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')} • v8")
