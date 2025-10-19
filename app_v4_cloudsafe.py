
import streamlit as st
import pandas as pd
import inspect

st.set_page_config(page_title="BHI PO → Invoices → Items", page_icon="📑", layout="wide")

# Determine which dataframe width API is supported
try:
    _DATAFRAME_SUPPORTS_WIDTH = "width" in inspect.signature(st.dataframe).parameters
except Exception:
    _DATAFRAME_SUPPORTS_WIDTH = False

def show_df(df, *, stretch: bool = True):
    \"\"\"Render a dataframe using the best width API available on this Streamlit version.\"\"\"
    if _DATAFRAME_SUPPORTS_WIDTH:
        return st.dataframe(df, width="stretch" if stretch else "content")
    else:
        return st.dataframe(df, use_container_width=True if stretch else False)

# Cache compatibility
if hasattr(st, "cache_data"):
    cache_like = st.cache_data
else:
    cache_like = st.cache

@cache_like
def load_workbook(default_path: str = "BHI.xlsx", uploaded_file=None):
    try:
        if uploaded_file is not None:
            xl = pd.ExcelFile(uploaded_file)
        else:
            xl = pd.ExcelFile(default_path)
        dfs = {s: xl.parse(s, dtype=object) for s in xl.sheet_names}
        for k, df in dfs.items():
            df.columns = [str(c).strip() for c in df.columns]
        return dfs, None
    except Exception as e:
        return {}, f"Failed to read workbook: {e}"

st.title("📑 PO → Invoices → Items (Cloud-Safe)")
st.caption("Type a PO number to see its invoices, then pick an invoice to see the line items. Compatible with old/new Streamlit versions.")

uploaded = st.file_uploader("Upload a .xlsx (optional). If omitted, the app will load BHI.xlsx from the same folder.", type=["xlsx"])
dfs, load_err = load_workbook(uploaded_file=uploaded)

if load_err:
    st.error(load_err)
    st.stop()
if not dfs:
    st.warning("No sheets found. Ensure BHI.xlsx is in the same folder or upload a file.")
    st.stop()

# Identify sheets
sheet_pos = "POs" if "POs" in dfs else list(dfs.keys())[0]
sheet_inv = "Invoices" if "Invoices" in dfs else (list(dfs.keys())[1] if len(dfs) > 1 else sheet_pos)
sheet_items = "InvoiceItems" if "InvoiceItems" in dfs else None

POs = dfs.get(sheet_pos, pd.DataFrame()).copy()
Invoices = dfs.get(sheet_inv, pd.DataFrame()).copy()
InvoiceItems = dfs.get(sheet_items, pd.DataFrame()) if sheet_items else pd.DataFrame()

# Normalize important fields
for df in (POs, Invoices, InvoiceItems):
    df.columns = [str(c).strip() for c in df.columns]

if "PO_NUMBER" in POs.columns:
    POs["PO_NUMBER"] = POs["PO_NUMBER"].astype(str).str.strip()
if "PO_NUMBER" in Invoices.columns:
    Invoices["PO_NUMBER"] = Invoices["PO_NUMBER"].astype(str).str.strip()
if "INVOICE_NUMBER" in Invoices.columns:
    Invoices["INVOICE_NUMBER"] = Invoices["INVOICE_NUMBER"].astype(str).str.strip()
if "INVOICE_NUMBER" in InvoiceItems.columns:
    InvoiceItems["INVOICE_NUMBER"] = InvoiceItems["INVOICE_NUMBER"].astype(str).str.strip()

# UI — PO input
po_input = st.text_input("Enter PO Number", placeholder="e.g., 4513194602").strip()

if po_input == "":
    st.info("Type a PO number above to see related invoices.")
    st.stop()

# Filter invoices by PO
if "PO_NUMBER" not in Invoices.columns:
    st.error("Invoices sheet missing 'PO_NUMBER' column.")
    st.stop()

inv_for_po = Invoices[Invoices["PO_NUMBER"].astype(str).str.fullmatch(po_input, case=False, na=False)]
if inv_for_po.empty:
    st.warning(f"No invoices found for PO: {po_input}")
    st.stop()

# Show invoice list for that PO
st.subheader("Invoices for this PO")
inv_cols = [c for c in ["INVOICE_NUMBER","INVOICE_DATE","INVOICE_AMOUNT","CURRENCY","STATUS","PO_NUMBER"] if c in inv_for_po.columns]
show_df(inv_for_po[inv_cols] if inv_cols else inv_for_po, stretch=True)

# Quick totals and coverage
def _num(x):
    try:
        return pd.to_numeric(str(x).replace(",","").replace(" ",""), errors="coerce")
    except Exception:
        return pd.NA

if "INVOICE_AMOUNT" in inv_for_po.columns:
    total_invoiced = _num(inv_for_po["INVOICE_AMOUNT"]).sum(skipna=True)
else:
    total_invoiced = pd.NA

if "PO_NUMBER" in POs.columns and "PO_AMOUNT" in POs.columns:
    po_row = POs[POs["PO_NUMBER"].astype(str) == po_input]
    if not po_row.empty:
        po_amount = _num(po_row.iloc[0]["PO_AMOUNT"])
        st.metric("Total Invoiced (this PO)", f"{(float(total_invoiced) if pd.notna(total_invoiced) else 0):,.0f}")
        if pd.notna(po_amount):
            st.metric("PO Amount", f"{float(po_amount):,.0f}")
            st.metric("Variance", f"{(float(po_amount) - (float(total_invoiced) if pd.notna(total_invoiced) else 0)):,.0f}")

# Select an invoice to see items
st.subheader("Pick an invoice to view its items")
invoice_numbers = inv_for_po["INVOICE_NUMBER"].dropna().astype(str).unique().tolist()
default_invoice = invoice_numbers[0] if invoice_numbers else None
picked_invoice = st.selectbox("Invoice Number", options=invoice_numbers, index=0 if default_invoice else None)

if picked_invoice is None:
    st.info("Select an invoice above to display its items.")
else:
    if InvoiceItems.empty:
        st.warning("No 'InvoiceItems' sheet found or it is empty.")
    else:
        items = InvoiceItems[InvoiceItems["INVOICE_NUMBER"].astype(str) == str(picked_invoice)]
        if items.empty:
            st.warning(f"No items found for invoice: {picked_invoice}")
        else:
            show_cols = [c for c in ["INVOICE_NUMBER","LINE","MATERIAL","DESCRIPTION","QTY","UNIT","UNIT_PRICE","LINE_TOTAL"] if c in items.columns]
            show_df(items[show_cols] if show_cols else items, stretch=True)
            st.download_button(
                "⬇️ Download items (CSV)",
                data=(items[show_cols] if show_cols else items).to_csv(index=False).encode("utf-8-sig"),
                file_name=f"{picked_invoice}_items.csv",
                mime="text/csv"
            )
