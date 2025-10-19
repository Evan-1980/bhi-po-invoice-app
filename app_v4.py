
import streamlit as st
import pandas as pd
import io
import hashlib
from io import BytesIO

st.set_page_config(page_title="BHI PO → Invoices → Items (v5)", page_icon="📑", layout="wide")

# -------------------- Helpers --------------------
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

# -------------------- UI: File --------------------
st.title("📑 PO → Invoices → Items (v5)")
st.caption("Type or search a PO, view its invoices, drill into line items, export CSV/Excel, and check data quality.")

uploaded = st.file_uploader("Upload a .xlsx (optional). If omitted, the app will load BHI.xlsx from the app folder.", type=["xlsx"])
dfs = get_dfs(uploaded)

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

REQUIRED = {
    "POs": ["PO_NUMBER", "PO_AMOUNT"],
    "Invoices": ["INVOICE_NUMBER", "PO_NUMBER", "INVOICE_AMOUNT"],
    "InvoiceItems": ["INVOICE_NUMBER"]
}
def warn_required(name, df):
    need = [c for c in REQUIRED[name] if c not in df.columns]
    if need:
        st.warning(f"'{name}' missing likely columns: {', '.join(need)}")

warn_required("POs", POs)
warn_required("Invoices", Invoices)
if not InvoiceItems.empty:
    warn_required("InvoiceItems", InvoiceItems)

# -------------------- PO search & select --------------------
left, right = st.columns([2, 3])
with left:
    po_query = st.text_input("Find PO (partial or full)", placeholder="e.g., 9460").strip()

if "PO_NUMBER" not in POs.columns or POs.empty:
    st.error("POs sheet is missing or has no 'PO_NUMBER'.")
    st.stop()

po_list = POs["PO_NUMBER"].dropna().astype(str).unique().tolist()
matches = [p for p in po_list if po_query.lower() in p.lower()] if po_query else po_list
matches = matches[:500]
picked_po = st.selectbox("Select PO", options=matches, index=0 if matches else None)

if not picked_po:
    st.info("Type to search or pick a PO.")
    st.stop()

# -------------------- Invoices for selected PO --------------------
if "PO_NUMBER" not in Invoices.columns:
    st.error("Invoices sheet missing 'PO_NUMBER'."); st.stop()

inv_for_po = Invoices[Invoices["PO_NUMBER"].astype(str) == str(picked_po)].copy()

st.subheader("Invoices for this PO")
inv_cols_pref = ["INVOICE_NUMBER","INVOICE_DATE","INVOICE_AMOUNT","CURRENCY","STATUS","PO_NUMBER"]
inv_cols = [c for c in inv_cols_pref if c in inv_for_po.columns]
st.dataframe(inv_for_po[inv_cols] if inv_cols else inv_for_po, use_container_width=True)

amt_ser = _num_series(inv_for_po["INVOICE_AMOUNT"]) if "INVOICE_AMOUNT" in inv_for_po.columns else pd.Series(dtype="float64")
total_invoiced = float(amt_ser.sum()) if not amt_ser.empty else 0.0

po_amount = None
if {"PO_NUMBER","PO_AMOUNT"}.issubset(POs.columns):
    r = POs[POs["PO_NUMBER"].astype(str) == str(picked_po)]
    if not r.empty:
        po_amount = float(_num_series(pd.Series([r.iloc[0]["PO_AMOUNT"]])).fillna(0).iloc[0])

c1, c2, c3 = st.columns(3)
c1.metric("Total Invoiced (this PO)", f"{total_invoiced:,.0f}")
if po_amount is not None:
    c2.metric("PO Amount", f"{po_amount:,.0f}")
    c3.metric("Variance", f"{(po_amount - total_invoiced):,.0f}")

# -------------------- All PO items (across invoices) --------------------
show_all_items = st.checkbox("Show all items for this PO", value=False)
po_items = pd.DataFrame()
if show_all_items:
    if InvoiceItems.empty or "INVOICE_NUMBER" not in InvoiceItems.columns:
        st.warning("No 'InvoiceItems' sheet found or it lacks 'INVOICE_NUMBER'.")
    else:
        inv_nums = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist() if "INVOICE_NUMBER" in inv_for_po.columns else []
        po_items = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str).isin(inv_nums)].copy()
        st.markdown("**All items across invoices for this PO**")
        items_pref = ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"]
        show_cols = [c for c in items_pref if c in po_items.columns]
        st.dataframe(po_items[show_cols] if show_cols else po_items, use_container_width=True)
        st.download_button(
            "⬇️ Download PO Items (CSV)",
            data=(po_items[show_cols] if show_cols else po_items).to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{picked_po}_all_items.csv",
            mime="text/csv"
        )

# -------------------- Pick invoice → items --------------------
st.subheader("Pick an invoice to view its items")
invoice_numbers = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist() if "INVOICE_NUMBER" in inv_for_po.columns else []
picked_invoice = st.selectbox("Invoice Number", options=invoice_numbers, index=0 if invoice_numbers else None)

if picked_invoice and not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns:
    items = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str) == str(picked_invoice)].copy()
    if items.empty:
        st.info("No items for the selected invoice.")
    else:
        items_pref = ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"]
        show_cols = [c for c in items_pref if c in items.columns]
        st.dataframe(items[show_cols] if show_cols else items, use_container_width=True)
        st.download_button(
            "⬇️ Download items (CSV)",
            data=(items[show_cols] if show_cols else items).to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{picked_invoice}_items.csv",
            mime="text/csv"
        )
else:
    st.info("Select an invoice above to display its items.")

# -------------------- Invoice totals table --------------------
if "INVOICE_NUMBER" in inv_for_po.columns:
    summary = (inv_for_po.assign(_amt=_num_series(inv_for_po.get("INVOICE_AMOUNT", pd.Series(dtype=str))))
               .groupby("INVOICE_NUMBER", dropna=True)["_amt"].sum().reset_index()
               .rename(columns={"_amt": "Invoice Total"}))
    st.markdown("**Invoice totals**")
    st.dataframe(summary, use_container_width=True)

# -------------------- Excel Pack Download (Invoices + Items) --------------------
def build_po_pack(po_num: str, inv_df: pd.DataFrame, items_df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        inv_df.to_excel(xw, index=False, sheet_name="Invoices")
        (items_df if not items_df.empty else pd.DataFrame()).to_excel(xw, index=False, sheet_name="Items")
    buf.seek(0)
    return buf

if "INVOICE_NUMBER" in inv_for_po.columns:
    if po_items.empty and not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns:
        inv_nums = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist()
        po_items = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str).isin(inv_nums)].copy()
    xbuf = build_po_pack(str(picked_po), inv_for_po, po_items)
    st.download_button("⬇️ Download PO Pack (Excel)", data=xbuf, file_name=f"{picked_po}_pack.xlsx")

# -------------------- Data Quality --------------------
st.markdown("---")
st.subheader("Data Quality Checks")
issues = []

if "INVOICE_NUMBER" in Invoices.columns:
    dup_n = int((Invoices["INVOICE_NUMBER"].astype(str).value_counts(dropna=False) > 1).sum())
    if dup_n:
        issues.append(f"Duplicate INVOICE_NUMBERs: {dup_n}")

if "PO_NUMBER" in Invoices.columns:
    miss_po = int(Invoices["PO_NUMBER"].isna().sum())
    if miss_po:
        issues.append(f"Invoices missing PO_NUMBER: {miss_po}")

if {"PO_NUMBER"}.issubset(Invoices.columns) and "PO_NUMBER" in POs.columns:
    known_pos = set(POs["PO_NUMBER"].dropna().astype(str))
    orphan = int((~Invoices["PO_NUMBER"].dropna().astype(str).isin(known_pos)).sum())
    if orphan:
        issues.append(f"Invoices referencing unknown PO_NUMBER: {orphan}")

if not InvoiceItems.empty and "INVOICE_NUMBER" in InvoiceItems.columns:
    miss_inv = int(InvoiceItems["INVOICE_NUMBER"].isna().sum())
    if miss_inv:
        issues.append(f"Items missing INVOICE_NUMBER: {miss_inv}")

if issues:
    for m in issues:
        st.warning("• " + m)
else:
    st.success("No obvious issues found.")
