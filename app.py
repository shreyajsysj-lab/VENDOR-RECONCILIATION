import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Ledger Reconciliation",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={'Get Help': None, 'Report a bug': None, 'About': None},
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

#MainMenu {visibility: hidden !important;}
footer {visibility: hidden !important;}
header {visibility: hidden !important;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
.stDeployButton {display: none !important;}

:root {
    --bg:           #F5F7FA;
    --surface:      #FFFFFF;
    --surface2:     #EFF2F7;
    --border:       #D1D9E6;
    --accent:       #1A3A6B;
    --accent2:      #1A6B45;
    --warn:         #B85C00;
    --danger:       #A32035;
    --text:         #1A202C;
    --muted:        #5A6A85;
    --matched:      #E8F5EE;
    --matched-border:#A8D5BA;
    --unmatched:    #FDE8EB;
    --unmatched-border:#F5A8B4;
    --partial:      #FFF3E0;
    --partial-border:#FFB74D;
    --vl-color:     #1A3A6B;
    --vl-bg:        #EEF3FF;
    --vl-border:    #A8BFEE;
    --cl-color:     #1A6B45;
    --cl-bg:        #EEF8F3;
    --cl-border:    #A8D5BA;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: var(--bg) !important;
    color: var(--text) !important;
}
.main { background: var(--bg) !important; }
.block-container { padding: 1.5rem 2rem !important; max-width: 1400px; }

[data-testid="stSidebar"] { background: #1A3A6B !important; border-right: none; }
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] .stMarkdown p { color: #CBD5E0 !important; }
[data-testid="stSidebar"] input { background: #2A4A8B !important; color: white !important; border-color: #4A6AAB !important; }

.recon-header {
    display: flex; align-items: center; gap: 1rem;
    margin-bottom: 1.5rem; padding: 1.25rem 1.5rem;
    border-radius: 12px;
    background: linear-gradient(135deg, #1A3A6B, #1A6B45);
}
.recon-logo { font-family: 'Poppins', sans-serif; font-size: 1.8rem; font-weight: 800; color: #FFFFFF; letter-spacing: -0.02em; }
.recon-subtitle { font-size: 0.72rem; color: #CBD5E0; text-transform: uppercase; letter-spacing: 0.12em; margin-top: 2px; }

.stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.5rem; }
.stat-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 12px; padding: 1.1rem 1.25rem;
    position: relative; overflow: hidden;
    box-shadow: 0 2px 8px rgba(26,58,107,0.07);
}
.stat-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; }
.stat-card.matched::before { background: #1A6B45; }
.stat-card.unmatched::before { background: #A32035; }
.stat-card.partial::before { background: #B85C00; }
.stat-card.total::before { background: #1A3A6B; }
.stat-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600; }
.stat-value { font-family: 'Poppins', sans-serif; font-size: 1.9rem; font-weight: 700; margin: 0.2rem 0; color: var(--text); }
.stat-card.matched .stat-value  { color: #1A6B45; }
.stat-card.unmatched .stat-value{ color: #A32035; }
.stat-card.partial .stat-value  { color: #B85C00; }
.stat-card.total .stat-value    { color: #1A3A6B; }
.stat-sub { font-size: 0.68rem; color: var(--muted); }

.stTabs [data-baseweb="tab-list"] { background: var(--surface2) !important; border-radius: 8px; padding: 4px; gap: 2px; border: 1px solid var(--border); }
.stTabs [data-baseweb="tab"] { font-family: 'Inter', sans-serif !important; font-size: 0.8rem !important; font-weight: 600 !important; color: var(--muted) !important; background: transparent !important; border-radius: 6px !important; padding: 6px 16px !important; }
.stTabs [aria-selected="true"] { background: #1A3A6B !important; color: #FFFFFF !important; }

[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; box-shadow: 0 1px 6px rgba(0,0,0,0.06); }

.stButton > button {
    font-family: 'Inter', sans-serif !important; font-weight: 600 !important;
    background: linear-gradient(135deg, #1A3A6B, #2A5AB0) !important;
    color: white !important; border: none !important; border-radius: 8px !important;
    padding: 0.55rem 1.4rem !important; transition: all 0.2s !important;
    box-shadow: 0 2px 8px rgba(26,58,107,0.25) !important;
}
.stButton > button:hover { opacity: 0.9 !important; transform: translateY(-1px) !important; }

.stDownloadButton > button {
    font-family: 'Inter', sans-serif !important; font-weight: 600 !important;
    background: var(--surface) !important; color: #1A6B45 !important;
    border: 1px solid #1A6B45 !important; border-radius: 8px !important;
}

[data-testid="stFileUploader"] { background: var(--surface) !important; border: 2px dashed var(--border) !important; border-radius: 10px !important; }

.section-tag { display: inline-block; font-family: 'Inter', sans-serif; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.12em; padding: 4px 10px; border-radius: 4px; margin-bottom: 0.5rem; }
.tag-matched   { background: var(--matched);   color: #1A6B45; border: 1px solid var(--matched-border); }
.tag-unmatched { background: var(--unmatched); color: #A32035; border: 1px solid var(--unmatched-border); }
.tag-partial   { background: var(--partial);   color: #B85C00; border: 1px solid var(--partial-border); }
.tag-blue      { background: var(--vl-bg);     color: #1A3A6B; border: 1px solid var(--vl-border); }
.tag-vl        { background: var(--vl-bg);     color: var(--vl-color); border: 1px solid var(--vl-border); }
.tag-cl        { background: var(--cl-bg);     color: var(--cl-color); border: 1px solid var(--cl-border); }

.info-box { background: var(--surface); border: 1px solid var(--border); border-left: 4px solid #1A3A6B; border-radius: 6px; padding: 0.85rem 1rem; font-size: 0.82rem; color: var(--text); margin-bottom: 1rem; line-height: 1.6; }

[data-testid="stAlert"] { border-radius: 8px !important; }
[data-testid="stNumberInput"] input { background: var(--surface) !important; border-color: var(--border) !important; color: var(--text) !important; border-radius: 8px !important; }
[data-testid="stExpander"] { background: var(--surface) !important; border: 1px solid var(--border) !important; border-radius: 10px !important; }
[data-testid="stTextInput"] input { background: var(--surface) !important; border-color: var(--border) !important; color: var(--text) !important; border-radius: 8px !important; font-family: 'Inter', sans-serif !important; }

.ai-badge { display: inline-flex; align-items: center; gap: 5px; background: linear-gradient(135deg,#1A3A6B,#1A6B45); color:#fff; font-size:0.65rem; font-weight:700; padding:3px 9px; border-radius:20px; letter-spacing:0.08em; }
.col-map-row { display:grid; grid-template-columns:1fr 1fr 80px; gap:8px; align-items:center; padding:6px 10px; border-bottom:1px solid var(--border); font-size:0.82rem; }
.col-map-row:nth-child(even) { background:var(--surface2); }
.col-field-name { font-weight:600; color:#1A3A6B; }
.col-detected   { color:#1A6B45; font-family:monospace; }
.col-confidence-high { color:#1A6B45; font-weight:700; }
.col-confidence-low  { color:#B85C00; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────

def fmt_inr(val):
    try:
        v = float(val)
        if v < 0:
            return f"(₹{abs(v):,.2f})"
        return f"₹{v:,.2f}"
    except:
        return "₹0.00"

def safe_sum(lst, key):
    try:
        return sum(float(d.get(key, 0) or 0) for d in lst)
    except:
        return 0.0

def display_df(data):
    if not data:
        st.caption("No records.")
        return
    df = pd.DataFrame(data)
    for col in df.columns:
        if 'date' in col.lower() or col in ['VL Date', 'CL Date']:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%b-%Y').fillna('')
    st.dataframe(df, use_container_width=True, hide_index=True)

def clean_doc_number(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    s = re.sub(r'[\s\-_/]', '', s)
    return s

def extract_utr(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    utr_pattern = re.search(r'[A-Z]{4}\d{18}|UTR[\s:]*([A-Z0-9]+)', s)
    if utr_pattern:
        return utr_pattern.group(0).replace(' ', '').replace(':', '')
    return s

def get_period(dt):
    try:
        return pd.to_datetime(dt).strftime('%Y-%m')
    except:
        return ""

def round_amount(val, decimals=2):
    try:
        return round(float(val), decimals)
    except:
        return 0.0

def is_debit_note(doc_type):
    if pd.isna(doc_type):
        return False
    s = str(doc_type).upper()
    return any(k in s for k in ['DEBIT NOTE', 'DN', 'DEBIT MEMO'])

def is_reversal_type(doc_type):
    if pd.isna(doc_type):
        return False
    s = str(doc_type).upper()
    return 'COMPLETE REVERSAL' in s

def is_credit_note(doc_type):
    if pd.isna(doc_type):
        return False
    s = str(doc_type).upper()
    return any(k in s for k in [
        'CREDIT NOTE', 'CREDIT MEMO', 'CREDIT',
        'SALEABLE RETURN', 'NON SALEABLE', 'NON-SALEABLE',
        'NONSALEABLE', 'SALE RETURN', 'SALES RETURN',
    ])

def is_discount_or_prn(doc_type, doc_no=""):
    if pd.isna(doc_type):
        doc_type = ""
    s = str(doc_type).upper()
    d = str(doc_no).upper()
    return any(k in s or k in d for k in [
        'DISCOUNT', 'DISC', 'DEBIT NOTE', 'DN',
        'PRN', 'PRICE REVISION', 'PRICE ADJ', 'REBATE',
        'ALLOWANCE', 'SCHEME', 'CLAIM'
    ])

def is_collection(doc_type, particulars=""):
    if pd.isna(doc_type):
        doc_type = ""
    s = str(doc_type).upper() + " " + str(particulars).upper()
    return any(k in s for k in ['PAYMENT', 'RECEIPT', 'COLLECTION', 'NEFT', 'RTGS', 'IMPS', 'CHEQUE', 'CHQ', 'TDS', 'BANK', 'UTR'])

def extract_ref_from_particulars(particulars):
    if pd.isna(particulars):
        return ""
    s = str(particulars).strip().upper()
    s = re.sub(r'(REVERSAL OF|AGAINST|REF|REFERENCE|RETURN OF|CANCELLATION OF|REVERSED)\s*', '', s)
    s = re.sub(r'[\s\-_/]', '', s)
    return s.strip()


# ─────────────────────────────────────────────
# AI COLUMN DETECTION (Claude API)
# ─────────────────────────────────────────────

def ai_detect_columns(columns: list, sample_rows: list, ledger_type: str = "vendor") -> dict:
    """
    Use Claude API to intelligently detect which column maps to which field.
    Returns dict: {field_name: detected_column_name, ...}
    """
    import requests

    fields_desc = {
        "doc_date":    "Transaction/document date (date of invoice, payment, etc.)",
        "doc_no":      "Document/voucher number (invoice no, cheque no, ref no)",
        "doc_type":    "Type of transaction (Invoice, Payment, Debit Note, Credit Note, etc.)",
        "particulars": "Description/narration/remarks of the transaction",
        "debit":       "Debit amount column (Dr side)",
        "credit":      "Credit amount column (Cr side)",
        "closing":     "Closing/running balance column",
    }
    if ledger_type == "customer":
        # customer ledgers often don't have particulars
        fields_desc.pop("particulars", None)

    # Build a sample of the data as a table string
    sample_str = ""
    if sample_rows:
        sample_str = "\nFirst few data rows:\n"
        for row in sample_rows[:5]:
            row_str = " | ".join([f"{k}: {v}" for k, v in list(row.items())[:len(columns)]])
            sample_str += f"  {row_str}\n"

    prompt = f"""You are a financial data expert helping to map columns from an Indian ERP/accounting ledger export.

Ledger type: {ledger_type.upper()} LEDGER

Available columns in the uploaded file:
{json.dumps(columns, indent=2)}
{sample_str}

Map each of these required fields to the best matching column from the list above.
Required fields:
{json.dumps(fields_desc, indent=2)}

Rules:
- Each column can only be mapped to ONE field
- If no column matches a field, use null
- Prefer exact or partial keyword matches
- For debit/credit: look for (Dr)/(Cr) suffixes, "Debit (LC)", "Credit Amount", etc.
- For doc_no: look for "No.", "Voucher", "Ref", "Document no", "Bill No", "Invoice No"
- For doc_type: look for "Type", "Nature", "Voucher Type", "Transaction Type"
- For closing: look for "Balance", "Closing", "Running Balance" — NOT "Opening"
- Return ONLY valid JSON, no explanation

Return format (JSON only):
{{
  "doc_date": "column name or null",
  "doc_no": "column name or null",
  "doc_type": "column name or null",
  "particulars": "column name or null",
  "debit": "column name or null",
  "credit": "column name or null",
  "closing": "column name or null",
  "confidence": {{"doc_date": "high/medium/low", "doc_no": "high/medium/low", ...}}
}}"""

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=15
        )
        if resp.status_code == 200:
            data = resp.json()
            text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
            # Strip markdown code fences if present
            text = re.sub(r"```(?:json)?", "", text).strip().rstrip("`").strip()
            result = json.loads(text)
            return result
    except Exception as e:
        st.warning(f"AI column detection unavailable ({e}). Using rule-based detection.")
    return {}


def _detect_header_row(df):
    HEADER_KEYWORDS = ['date', 'debit', 'credit', 'doc', 'document',
                       'voucher', 'vch', 'amount', 'balance', 'no.',
                       'number', 'type', 'particular', 'narration', 'ref',
                       'invoice', 'bill', 'transaction', 'nature']
    best_row = 0
    best_score = 0
    for i in range(min(25, len(df))):
        try:
            row_vals = df.iloc[i].tolist()
        except Exception:
            continue
        score = 0
        non_null = 0
        for v in row_vals:
            try:
                s = str(v).lower().strip()
                if s and s not in ['nan', 'none', '']:
                    non_null += 1
                if any(k in s for k in HEADER_KEYWORDS):
                    score += 1
            except Exception:
                continue
        # Weighted: keyword hits + density bonus
        weighted = score * 2 + (1 if non_null >= 3 else 0)
        if weighted > best_score:
            best_score = weighted
            best_row = i
    return best_row


def _get_closing_from_raw(df):
    try:
        for c in df.columns:
            col_name = str(c).lower().strip()
            if 'closing' in col_name or ('balance' in col_name and 'opening' not in col_name):
                series = pd.to_numeric(df[c], errors='coerce').dropna()
                if not series.empty:
                    return float(series.iloc[-1])
    except Exception:
        pass
    return None


def _rule_based_map_columns(df):
    """Fallback keyword-based column mapper."""
    col_map = {}
    already_mapped = set()

    def try_map(target, col_name):
        if target not in already_mapped:
            col_map[col_name] = target
            already_mapped.add(target)

    for c in df.columns:
        cl = str(c).lower().strip()

        if 'date' in cl and 'doc_date' not in already_mapped:
            try_map('doc_date', c)
        elif ('type' in cl or 'nature' in cl) and 'doc_type' not in already_mapped and 'date' not in cl:
            try_map('doc_type', c)
        elif 'doc_no' not in already_mapped and 'date' not in cl and 'type' not in cl \
                and any(k in cl for k in ['no.', 'no ', 'num', 'voucher', 'vch', 'ref', 'detail', 'invoice no', 'bill no', 'document no']):
            try_map('doc_no', c)
        elif any(k in cl for k in ['particular', 'narration', 'description', 'remarks', 'remark']) \
                and 'particulars' not in already_mapped:
            try_map('particulars', c)
        elif 'opening' in cl:
            pass  # skip
        elif 'debit' in cl and 'credit' not in cl and 'debit' not in already_mapped:
            try_map('debit', c)
        elif 'credit' in cl and 'debit' not in cl and 'credit' not in already_mapped:
            try_map('credit', c)
        elif ('closing' in cl or 'balance' in cl) and 'opening' not in cl and 'closing' not in already_mapped:
            try_map('closing', c)

    return col_map


def _apply_column_mapping(df, mapping: dict) -> pd.DataFrame:
    """
    Given a mapping {standard_field: original_col_name}, rename and ensure all fields exist.
    mapping may come from AI or manual selection.
    """
    # Reverse: original_col -> standard_field
    rename_dict = {}
    for field, orig_col in mapping.items():
        if orig_col and orig_col in df.columns and field not in ['confidence']:
            rename_dict[orig_col] = field

    df = df.rename(columns=rename_dict)
    return df


def _load_any_ledger_smart(file, is_vendor=True, override_mapping=None):
    """
    Universal ledger loader with AI column detection.
    Returns (df, closing_val, detected_mapping, raw_columns)
    """
    try:
        raw = pd.read_excel(file, header=None, dtype=str)
    except Exception:
        raw = pd.read_excel(file, header=None)

    header_idx = _detect_header_row(raw)

    try:
        hdr_series = raw.iloc[header_idx]
        col_names = []
        for i, v in enumerate(hdr_series):
            try:
                cell = str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else ''
                col_names.append(cell if cell else f'_Col{i}')
            except Exception:
                col_names.append(f'_Col{i}')
    except Exception:
        col_names = [f'_Col{i}' for i in range(len(raw.columns))]

    data = raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    data.columns = col_names
    raw_columns = list(col_names)

    closing_val = _get_closing_from_raw(data)

    # Build sample rows for AI
    sample_rows = []
    for _, row in data.head(8).iterrows():
        sample_rows.append({col: str(row[col]) for col in col_names if str(row[col]).strip() not in ['', 'nan', 'None']})

    # Column mapping: override > AI > rule-based
    if override_mapping:
        final_mapping = override_mapping
    else:
        ai_result = ai_detect_columns(raw_columns, sample_rows, "vendor" if is_vendor else "customer")
        if ai_result and any(v for k, v in ai_result.items() if k != 'confidence' and v):
            final_mapping = {k: v for k, v in ai_result.items() if k != 'confidence' and v and v in raw_columns}
        else:
            # Fallback: rule-based
            rule_map = _rule_based_map_columns(data)
            final_mapping = {v: k for k, v in rule_map.items()}  # field -> orig_col

    confidence = ai_result.get('confidence', {}) if 'ai_result' in dir() else {}

    df = _apply_column_mapping(data.copy(), final_mapping)

    needed = ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit', 'closing']
    if is_vendor:
        needed += ['particulars']
    for col in needed:
        if col not in df.columns:
            df[col] = ''

    try:
        df['doc_date'] = pd.to_datetime(df['doc_date'], errors='coerce', dayfirst=True)
    except Exception:
        df['doc_date'] = pd.NaT

    for num_col, fill in [('debit', 0), ('credit', 0)]:
        try:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce').fillna(fill)
        except Exception:
            df[num_col] = fill

    try:
        df['closing'] = pd.to_numeric(df['closing'], errors='coerce')
    except Exception:
        df['closing'] = np.nan

    try:
        df['doc_no'] = df['doc_no'].fillna('').astype(str).str.strip()
        df['doc_no'] = df['doc_no'].replace({'nan': '', 'None': '', 'NaN': ''})
    except Exception:
        df['doc_no'] = ''

    df['doc_no_clean'] = df['doc_no'].apply(clean_doc_number)
    df['period'] = df['doc_date'].apply(get_period)

    if is_vendor:
        try:
            df['particulars'] = df['particulars'].fillna('').astype(str)
        except Exception:
            df['particulars'] = ''
        df['particulars_ref'] = df['particulars'].apply(extract_ref_from_particulars)

    df = df[df['doc_no_clean'].astype(str) != ''].reset_index(drop=True)
    df['_idx'] = df.index
    df['_remark'] = ''
    df['_match_ref'] = ''

    return df, closing_val, final_mapping, raw_columns, confidence


def load_vendor_ledger(file, override_mapping=None):
    df, closing, mapping, raw_cols, confidence = _load_any_ledger_smart(file, is_vendor=True, override_mapping=override_mapping)
    df._vl_closing = closing
    return df, closing, mapping, raw_cols, confidence


def load_customer_ledger(file, override_mapping=None):
    df, closing, mapping, raw_cols, confidence = _load_any_ledger_smart(file, is_vendor=False, override_mapping=override_mapping)
    df._cl_closing = closing
    return df, closing, mapping, raw_cols, confidence


# ─────────────────────────────────────────────
# RECONCILIATION ENGINE
# ─────────────────────────────────────────────

def run_reconciliation(vl_orig, cl_orig, tolerance=1.0):
    results = {
        'invoice_matched': [],
        'invoice_unmatched_vl': [],
        'invoice_unmatched_cl': [],
        'cn_unmatched_vl': [],
        'dn_matched': [],
        'dn_unmatched_vl': [],
        'dn_unmatched_cl': [],
        'collection_matched': [],
        'collection_unmatched_vl': [],
        'collection_unmatched_cl': [],
        'reversal_vl_internal': [],
        'reversal_cross_ledger': [],
        'reversal_unmatched': [],
    }

    vl = vl_orig.copy()
    cl = cl_orig.copy()
    vl['_matched'] = False
    cl['_matched'] = False

    # ═══ STEP 1: Process VL Reversal entries ═══
    vl_reversals = vl[vl['doc_type'].apply(is_reversal_type)].copy()

    def get_vl_invoice_pool(vl_df):
        return vl_df[
            (~vl_df['_matched']) &
            (~vl_df['doc_type'].apply(is_reversal_type)) &
            (~vl_df['doc_type'].apply(is_debit_note)) &
            (~vl_df['doc_type'].apply(lambda x: is_collection(x)))
        ]

    for idx, rev_row in vl_reversals.iterrows():
        ref_particulars = rev_row.get('particulars_ref', '')
        raw_particulars = str(rev_row.get('particulars', ''))
        rev_amount = round_amount(rev_row.get('debit', 0) + rev_row.get('credit', 0))

        orig_pool = get_vl_invoice_pool(vl)
        orig_match = None
        match_basis_rev = ''

        if ref_particulars:
            m = orig_pool[orig_pool['doc_no_clean'] == ref_particulars]
            if not m.empty:
                orig_match = m.iloc[0]
                match_basis_rev = 'Particulars Reference (Exact)'

        if orig_match is None:
            words = re.findall(r'[A-Z0-9]{4,}', raw_particulars.upper())
            for word in words:
                m = orig_pool[orig_pool['doc_no_clean'] == word]
                if not m.empty:
                    orig_match = m.iloc[0]
                    match_basis_rev = f'Particulars Word Match ({word})'
                    break

        if orig_match is None and ref_particulars and len(ref_particulars) >= 5:
            prefix = ref_particulars[:8]
            m = orig_pool[orig_pool['doc_no_clean'].str.startswith(prefix, na=False)]
            if not m.empty:
                orig_match = m.iloc[0]
                match_basis_rev = 'Particulars Partial Match'

        if orig_match is None and rev_amount > 0:
            m = orig_pool[
                (orig_pool['period'] == rev_row.get('period', '')) &
                (abs(orig_pool['debit'] + orig_pool['credit'] - rev_amount) <= tolerance)
            ]
            if not m.empty:
                orig_match = m.iloc[0]
                match_basis_rev = 'Period + Amount Match'

        amount_valid = False
        if orig_match is not None:
            orig_amount = round_amount(orig_match.get('debit', 0) + orig_match.get('credit', 0))
            if orig_amount > 0 and abs(rev_amount - orig_amount) <= tolerance:
                amount_valid = True

        if orig_match is not None and amount_valid:
            orig_vl_idx = orig_match['_idx']
            orig_doc_no = str(orig_match.get('doc_no', ''))
            orig_amount_val = round_amount(orig_match.get('debit', 0) + orig_match.get('credit', 0))

            cl_pool = cl[
                (~cl['_matched']) &
                (~cl['doc_type'].apply(is_debit_note)) &
                (~cl['doc_type'].apply(lambda x: is_collection(x, '')))
            ]
            cl_for_orig = cl_pool[cl_pool['doc_no_clean'] == orig_match['doc_no_clean']]

            vl.at[idx, '_matched'] = True
            vl.at[idx, '_match_ref'] = orig_doc_no
            vl.at[orig_vl_idx, '_matched'] = True
            vl.at[orig_vl_idx, '_match_ref'] = str(rev_row.get('doc_no', ''))

            if not cl_for_orig.empty:
                cl_row = cl_for_orig.iloc[0]
                cl.at[cl_row['_idx'], '_matched'] = True
                cl.at[cl_row['_idx'], '_remark'] = 'Invoice Reversed in VL — Needs Review'
                cl.at[cl_row['_idx'], '_match_ref'] = orig_doc_no
                vl.at[idx, '_remark'] = 'Reversal Entry — Invoice Also in CL'
                vl.at[orig_vl_idx, '_remark'] = 'Invoice Reversed in VL — Also in CL'
                results['reversal_cross_ledger'].append({
                    'VL Original Doc No': orig_doc_no,
                    'VL Original Date': orig_match.get('doc_date', ''),
                    'VL Original Type': orig_match.get('doc_type', ''),
                    'VL Original Debit': orig_match.get('debit', 0),
                    'VL Original Credit': orig_match.get('credit', 0),
                    'VL Reversal Doc No': str(rev_row.get('doc_no', '')),
                    'VL Reversal Date': rev_row.get('doc_date', ''),
                    'VL Reversal Type': rev_row.get('doc_type', ''),
                    'VL Reversal Debit': rev_row.get('debit', 0),
                    'VL Reversal Credit': rev_row.get('credit', 0),
                    'VL Reversal Particulars': raw_particulars,
                    'CL Doc No': cl_row.get('doc_no', ''),
                    'CL Date': cl_row.get('doc_date', ''),
                    'CL Type': cl_row.get('doc_type', ''),
                    'CL Debit': cl_row.get('debit', 0),
                    'CL Credit': cl_row.get('credit', 0),
                    'Match Basis': match_basis_rev,
                    'Amount Match': f'VL={rev_amount} | Orig={orig_amount_val}',
                    'Remark': 'Invoice Reversed in VL but Present in CL — Needs Review',
                })
            else:
                vl.at[idx, '_remark'] = 'Reversal Entry — Not in CL'
                vl.at[orig_vl_idx, '_remark'] = 'Invoice Reversed in VL — Not in CL'
                results['reversal_vl_internal'].append({
                    'VL Original Doc No': orig_doc_no,
                    'VL Original Date': orig_match.get('doc_date', ''),
                    'VL Original Type': orig_match.get('doc_type', ''),
                    'VL Original Debit': orig_match.get('debit', 0),
                    'VL Original Credit': orig_match.get('credit', 0),
                    'VL Reversal Doc No': str(rev_row.get('doc_no', '')),
                    'VL Reversal Date': rev_row.get('doc_date', ''),
                    'VL Reversal Type': rev_row.get('doc_type', ''),
                    'VL Reversal Debit': rev_row.get('debit', 0),
                    'VL Reversal Credit': rev_row.get('credit', 0),
                    'VL Reversal Particulars': raw_particulars,
                    'Match Basis': match_basis_rev,
                    'Amount Match': f'VL={rev_amount} | Orig={orig_amount_val}',
                    'Remark': 'Invoice Reversed in VL — Not Present in CL',
                })
        else:
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = 'Reversal Entry — Original Not Found / Amount Mismatch'
            reason = 'Amount Mismatch' if orig_match is not None else 'Original Not Found'
            results['reversal_unmatched'].append({
                'VL Doc No': str(rev_row.get('doc_no', '')),
                'VL Date': rev_row.get('doc_date', ''),
                'VL Type': rev_row.get('doc_type', ''),
                'VL Debit': rev_row.get('debit', 0),
                'VL Credit': rev_row.get('credit', 0),
                'VL Reversal Particulars': raw_particulars,
                'Reason': reason,
                'Remark': f'Reversal — {reason}',
            })

    # ═══ STEP 2A: Match Invoices by Doc Number ═══
    vl_inv = vl[
        (~vl['_matched']) &
        (~vl['doc_type'].apply(is_debit_note)) &
        (~vl['doc_type'].apply(is_credit_note)) &
        (~vl['doc_type'].apply(is_collection)) &
        (~vl['doc_type'].apply(is_reversal_type))
    ].copy()

    cl_inv = cl[
        (~cl['_matched']) &
        (~cl['doc_type'].apply(is_debit_note)) &
        (~cl['doc_type'].apply(lambda x: is_collection(x, '')))
    ].copy()

    for idx, vrow in vl_inv.iterrows():
        doc_matches = cl_inv[
            (cl_inv['doc_no_clean'] == vrow['doc_no_clean']) &
            (~cl_inv['_matched']) &
            (vrow['doc_no_clean'] != '')
        ]
        if not doc_matches.empty:
            crow = doc_matches.iloc[0]
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = 'Matched — Invoice'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = 'Matched — Invoice'
            cl.at[crow['_idx'], '_match_ref'] = str(vrow.get('doc_no', ''))
            cl_inv.at[crow.name, '_matched'] = True
            results['invoice_matched'].append({
                'VL Doc No': vrow.get('doc_no', ''),
                'VL Date': vrow.get('doc_date', ''),
                'VL Type': vrow.get('doc_type', ''),
                'VL Debit': vrow.get('debit', 0),
                'VL Credit': vrow.get('credit', 0),
                'CL Doc No': crow.get('doc_no', ''),
                'CL Date': crow.get('doc_date', ''),
                'CL Type': crow.get('doc_type', ''),
                'CL Debit': crow.get('debit', 0),
                'CL Credit': crow.get('credit', 0),
                'Match Basis': 'Document Number',
                'Match Type': 'Invoice',
                'Remark': 'Matched — Invoice',
            })

    # ═══ STEP 2B: Match Credit Notes vs Discount/PRN ═══
    vl_cn = vl[(~vl['_matched']) & (vl['doc_type'].apply(is_credit_note))].copy()
    cl_disc = cl[
        (~cl['_matched']) &
        (cl.apply(lambda r: is_discount_or_prn(r.get('doc_type', ''), r.get('doc_no', '')), axis=1))
    ].copy()
    cl_any_unmatched = cl[(~cl['_matched']) & (~cl['doc_type'].apply(lambda x: is_collection(x, '')))].copy()

    for idx, vrow in vl_cn.iterrows():
        matched = False
        basis = ''
        crow = None

        if not cl_disc.empty:
            doc_m = cl_disc[(cl_disc['doc_no_clean'] == vrow['doc_no_clean']) & (~cl_disc['_matched'])]
            if not doc_m.empty:
                crow = doc_m.iloc[0]
                matched = True
                basis = 'Document Number (Credit Note ↔ Discount/PRN)'

        if not matched and not cl_disc.empty:
            vamt = round_amount(vrow['debit'] + vrow['credit'])
            if vamt > 0:
                amt_m = cl_disc[
                    (cl_disc['period'] == vrow['period']) &
                    (abs(cl_disc['debit'] + cl_disc['credit'] - vamt) <= tolerance) &
                    (~cl_disc['_matched'])
                ]
                if not amt_m.empty:
                    crow = amt_m.iloc[0]
                    matched = True
                    basis = 'Period + Amount (Credit Note ↔ Discount/PRN)'

        if not matched:
            doc_m2 = cl_any_unmatched[
                (cl_any_unmatched['doc_no_clean'] == vrow['doc_no_clean']) &
                (~cl_any_unmatched['_matched'])
            ]
            if not doc_m2.empty:
                crow = doc_m2.iloc[0]
                matched = True
                basis = 'Document Number (Credit Note ↔ CL Entry)'

        if matched and crow is not None:
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = f'Matched — Credit Note ({basis.split("(")[0].strip()})'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = f'Matched — Credit Note / Discount-PRN'
            cl.at[crow['_idx'], '_match_ref'] = str(vrow.get('doc_no', ''))
            if crow.name in cl_disc.index:
                cl_disc.at[crow.name, '_matched'] = True
            if crow.name in cl_any_unmatched.index:
                cl_any_unmatched.at[crow.name, '_matched'] = True
            results['dn_matched'].append({
                'VL Doc No': vrow.get('doc_no', ''),
                'VL Date': vrow.get('doc_date', ''),
                'VL Type': vrow.get('doc_type', ''),
                'VL Debit': vrow.get('debit', 0),
                'VL Credit': vrow.get('credit', 0),
                'CL Doc No': crow.get('doc_no', ''),
                'CL Date': crow.get('doc_date', ''),
                'CL Type': crow.get('doc_type', ''),
                'CL Debit': crow.get('debit', 0),
                'CL Credit': crow.get('credit', 0),
                'Match Basis': basis,
                'Match Type': 'Credit Note ↔ Discount / PRN',
                'Remark': f'Matched — Credit Note vs Discount/PRN',
            })

    # ═══ STEP 3: Debit Notes ═══
    vl_dn = vl[(~vl['_matched']) & (vl['doc_type'].apply(is_debit_note))].copy()
    cl_dn = cl[(~cl['_matched']) & (cl['doc_type'].apply(is_debit_note))].copy()

    for idx, vrow in vl_dn.iterrows():
        matched = False
        basis = ''
        doc_matches = cl_dn[(cl_dn['doc_no_clean'] == vrow['doc_no_clean']) & (~cl_dn['_matched'])]
        if not doc_matches.empty:
            crow = doc_matches.iloc[0]
            matched = True
            basis = 'Document Number'
        else:
            vamt = round_amount(vrow['debit'] + vrow['credit'])
            period_matches = cl_dn[
                (cl_dn['period'] == vrow['period']) &
                (abs(cl_dn['debit'] + cl_dn['credit'] - vamt) <= tolerance) &
                (~cl_dn['_matched'])
            ]
            if not period_matches.empty:
                crow = period_matches.iloc[0]
                matched = True
                basis = 'Period + Amount'

        if matched:
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = f'Matched — Debit Note ({basis})'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = f'Matched — Debit Note ({basis})'
            cl.at[crow['_idx'], '_match_ref'] = str(vrow.get('doc_no', ''))
            cl_dn.at[crow.name, '_matched'] = True
            results['dn_matched'].append({
                'VL Doc No': vrow.get('doc_no', ''),
                'VL Date': vrow.get('doc_date', ''),
                'VL Type': vrow.get('doc_type', ''),
                'VL Debit': vrow.get('debit', 0),
                'VL Credit': vrow.get('credit', 0),
                'CL Doc No': crow.get('doc_no', ''),
                'CL Date': crow.get('doc_date', ''),
                'CL Type': crow.get('doc_type', ''),
                'CL Debit': crow.get('debit', 0),
                'CL Credit': crow.get('credit', 0),
                'Match Basis': basis,
                'Match Type': 'Debit Note',
                'Remark': f'Matched — Debit Note ({basis})',
            })

    # ═══ STEP 4: Collections ═══
    vl_col = vl[(~vl['_matched']) & (vl['doc_type'].apply(lambda x: is_collection(x, '')))].copy()
    cl_col = cl[(~cl['_matched']) & (cl['doc_type'].apply(lambda x: is_collection(x, '')))].copy()

    vl_col['utr'] = vl_col.apply(lambda r: extract_utr(str(r.get('particulars', '')) + ' ' + str(r.get('doc_no', ''))), axis=1)
    cl_col['utr'] = cl_col.apply(lambda r: extract_utr(str(r.get('doc_no', '')) + ' ' + str(r.get('doc_type', ''))), axis=1)

    for idx, vrow in vl_col.iterrows():
        matched = False
        basis = ''
        crow = None
        if vrow['utr']:
            utr_matches = cl_col[(cl_col['utr'] == vrow['utr']) & (cl_col['utr'] != '') & (~cl_col['_matched'])]
            if not utr_matches.empty:
                crow = utr_matches.iloc[0]
                matched = True
                basis = 'UTR Number'

        if not matched:
            vamt = round_amount(vrow['debit'] + vrow['credit'])
            amt_matches = cl_col[
                (cl_col['period'] == vrow['period']) &
                (abs(cl_col['debit'] + cl_col['credit'] - vamt) <= tolerance) &
                (~cl_col['_matched'])
            ]
            if not amt_matches.empty:
                crow = amt_matches.iloc[0]
                matched = True
                basis = 'Period + Amount'

        if matched:
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = f'Matched — Collection ({basis})'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = f'Matched — Collection ({basis})'
            cl.at[crow['_idx'], '_match_ref'] = str(vrow.get('doc_no', ''))
            cl_col.at[crow.name, '_matched'] = True
            results['collection_matched'].append({
                'VL Doc No': vrow.get('doc_no', ''),
                'VL Date': vrow.get('doc_date', ''),
                'VL Type': vrow.get('doc_type', ''),
                'VL Amount': vrow.get('debit', 0) + vrow.get('credit', 0),
                'VL UTR': vrow.get('utr', ''),
                'CL Doc No': crow.get('doc_no', ''),
                'CL Date': crow.get('doc_date', ''),
                'CL Type': crow.get('doc_type', ''),
                'CL Amount': crow.get('debit', 0) + crow.get('credit', 0),
                'CL UTR': crow.get('utr', ''),
                'Match Basis': basis,
                'Match Type': 'Collection',
                'Remark': f'Matched — Collection ({basis})',
            })

    # ═══ STEP 5: Unmatched ═══
    for idx, r in vl[~vl['_matched']].iterrows():
        doc_t = str(r.get('doc_type', ''))
        if is_credit_note(doc_t):
            unmatched_remark = 'Unmatched — Credit Note'
        elif is_reversal_type(doc_t):
            unmatched_remark = 'Unmatched — Reversal Entry'
        else:
            unmatched_remark = 'Unmatched — Invoice'
        vl.at[idx, '_remark'] = unmatched_remark
        entry = {
            'Doc No': r.get('doc_no', ''),
            'Date': r.get('doc_date', ''),
            'Type': r.get('doc_type', ''),
            'Particulars': r.get('particulars', ''),
            'Debit': r.get('debit', 0),
            'Credit': r.get('credit', 0),
            'Source': 'Vendor Ledger',
            'Remark': unmatched_remark,
        }
        if is_debit_note(doc_t):
            results['dn_unmatched_vl'].append(entry)
        elif is_collection(doc_t):
            results['collection_unmatched_vl'].append(entry)
        elif is_credit_note(doc_t):
            results['cn_unmatched_vl'].append(entry)
        else:
            results['invoice_unmatched_vl'].append(entry)

    for idx, r in cl[~cl['_matched']].iterrows():
        doc_t = str(r.get('doc_type', ''))
        if is_credit_note(doc_t):
            cl_unmatched_remark = 'Unmatched — Credit Note'
        else:
            cl_unmatched_remark = 'Unmatched — Invoice'
        cl.at[idx, '_remark'] = cl_unmatched_remark
        entry = {
            'Doc No': r.get('doc_no', ''),
            'Date': r.get('doc_date', ''),
            'Type': r.get('doc_type', ''),
            'Debit': r.get('debit', 0),
            'Credit': r.get('credit', 0),
            'Source': 'Customer Ledger',
            'Remark': cl_unmatched_remark,
        }
        if is_debit_note(doc_t):
            results['dn_unmatched_cl'].append(entry)
        elif is_collection(doc_t):
            results['collection_unmatched_cl'].append(entry)
        else:
            results['invoice_unmatched_cl'].append(entry)

    results['vl_annotated'] = vl
    results['cl_annotated'] = cl
    return results


# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────

def build_excel(results, vl_orig, cl_orig, VL='Vendor', CL='Customer'):
    output = BytesIO()
    wb = openpyxl.Workbook()

    VL_COLOR = '1A3A6B'; CL_COLOR = '1A6B45'
    VL_LIGHT = 'D6E4FF'; CL_LIGHT = 'D6F5EA'
    MTH_COLOR = '1A6B45'; UNM_COLOR = 'A32035'
    MTH_FILL = 'C6EFCE'; UNM_FILL = 'FFC7CE'; REV_FILL = 'FFEB9C'; DARK = '1C2130'
    COLORS = {'matched_fill': MTH_FILL, 'unmatched_fill': UNM_FILL, 'reversal_fill': REV_FILL, 'border': 'C0C8D8'}

    thin = Side(style='thin', color=COLORS['border'])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def mk_fill(h): return PatternFill(fill_type='solid', fgColor=h)

    def style_header(ws, headers, row=1, color=DARK):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
            cell.fill = mk_fill(color)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        ws.row_dimensions[row].height = 30

    def auto_width(ws, min_w=10, max_w=48):
        for col in ws.columns:
            max_len = 0
            for cell in col:
                try: max_len = max(max_len, len(str(cell.value or '')))
                except: pass
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w, max(min_w, max_len + 2))

    def fmt_date(val):
        try:
            if pd.isna(val): return ''
            return pd.to_datetime(val).strftime('%d-%b-%Y')
        except: return str(val)

    def write_cell(ws, row, col, val):
        cell = ws.cell(row=row, column=col)
        if isinstance(val, (pd.Timestamp, datetime)): cell.value = fmt_date(val)
        elif isinstance(val, float) and not pd.isna(val):
            cell.value = round(val, 2); cell.number_format = '#,##0.00'
        elif not isinstance(val, str) and pd.isna(val): cell.value = ''
        else: cell.value = val
        return cell

    def ssum(lst, key):
        try: return sum(float(d.get(key, 0) or 0) for d in lst)
        except: return 0.0

    def pct(m, t): return f'{round(m/t*100,1)}%' if t else '0%'

    cn_list = [r for r in results['dn_matched'] if 'Credit Note' in str(r.get('Match Type', ''))]
    dn_list = [r for r in results['dn_matched'] if 'Credit Note' not in str(r.get('Match Type', ''))]

    # Write matched/unmatched sheets
    def write_sheet(wb, sheet_name, data, color=DARK):
        if not data:
            ws = wb.create_sheet(sheet_name[:31])
            ws['A1'] = 'No records'
            return ws
        ws = wb.create_sheet(sheet_name[:31])
        ws.sheet_view.showGridLines = False
        headers = list(data[0].keys())
        style_header(ws, headers, row=1, color=color)
        for ri, row_dict in enumerate(data, 2):
            remark = str(row_dict.get('Remark', ''))
            if 'Unmatched' in remark or 'Mismatch' in remark:
                fill = UNM_FILL
            elif 'Reversal' in remark:
                fill = REV_FILL
            else:
                fill = MTH_FILL
            for ci, key in enumerate(headers, 1):
                val = row_dict.get(key, '')
                cell = write_cell(ws, ri, ci, val)
                cell.fill = mk_fill(fill)
                cell.border = border
                cell.font = Font(name='Calibri', size=9)
                cell.alignment = Alignment(vertical='center')
            ws.row_dimensions[ri].height = 18
        auto_width(ws)
        return ws

    # Remove default sheet
    default_ws = wb.active
    default_ws.title = 'Summary'
    ws_sum = default_ws
    ws_sum.sheet_view.showGridLines = False

    # Summary sheet
    ws_sum['A1'] = f'⚖️ Ledger Reconciliation — {VL} vs {CL}'
    ws_sum['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Calibri')
    ws_sum['A1'].fill = mk_fill(DARK)
    ws_sum.merge_cells('A1:M1')
    ws_sum['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[1].height = 36

    sum_headers = ['Category', 'VL Count', 'VL Value', 'Matched Cnt', 'Matched Val',
                   'CL Count', 'CL Value', 'Unmatched Cnt', 'Unmatched Val', 'Match%', 'Remark']
    style_header(ws_sum, sum_headers, row=2, color='1A3A6B')

    def sum_row(ws, r, label, vl_c, vl_v, mc, mv, cl_c, cl_v, uc, uv, pct_str, rem, fill='F5F7FA'):
        vals = [label, vl_c, vl_v, mc, mv, cl_c, cl_v, uc, uv, pct_str, rem]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=ci, value=val)
            cell.fill = mk_fill(fill)
            cell.border = border
            cell.font = Font(name='Calibri', size=9)
            cell.alignment = Alignment(vertical='center', wrap_text=(ci in [1, 11]))
            if isinstance(val, float) and ci in [3, 5, 7, 9]:
                cell.number_format = '#,##0.00'
            elif isinstance(val, int) and ci in [2, 4, 6, 8]:
                cell.number_format = '#,##0'
        ws.row_dimensions[r].height = 20

    inv_vl_t = len(results['invoice_matched']) + len(results['invoice_unmatched_vl'])
    dn_vl_t  = len(dn_list) + len(results['dn_unmatched_vl'])
    col_vl_t = len(results['collection_matched']) + len(results['collection_unmatched_vl'])

    sum_row(ws_sum, 3, 'Invoices',
            inv_vl_t, ssum(results['invoice_matched'],'VL Debit')+ssum(results['invoice_unmatched_vl'],'Debit'),
            len(results['invoice_matched']), ssum(results['invoice_matched'],'VL Debit')+ssum(results['invoice_matched'],'VL Credit'),
            len(results['invoice_matched'])+len(results['invoice_unmatched_cl']),
            ssum(results['invoice_matched'],'CL Debit')+ssum(results['invoice_unmatched_cl'],'Debit'),
            len(results['invoice_unmatched_vl']),
            ssum(results['invoice_unmatched_vl'],'Debit')+ssum(results['invoice_unmatched_vl'],'Credit'),
            pct(len(results['invoice_matched']), inv_vl_t), 'Matched by Document Number', 'EEF3FF')

    sum_row(ws_sum, 4, 'Debit Notes',
            dn_vl_t, ssum(dn_list,'VL Debit')+ssum(results['dn_unmatched_vl'],'Debit'),
            len(dn_list), ssum(dn_list,'VL Debit')+ssum(dn_list,'VL Credit'),
            len(dn_list)+len(results['dn_unmatched_cl']),
            ssum(dn_list,'CL Debit')+ssum(results['dn_unmatched_cl'],'Debit'),
            len(results['dn_unmatched_vl']),
            ssum(results['dn_unmatched_vl'],'Debit')+ssum(results['dn_unmatched_vl'],'Credit'),
            pct(len(dn_list), dn_vl_t), 'Doc No → Period+Amount', 'EEF8F3')

    sum_row(ws_sum, 5, 'Collections',
            col_vl_t, ssum(results['collection_matched'],'VL Amount')+ssum(results['collection_unmatched_vl'],'Debit'),
            len(results['collection_matched']), ssum(results['collection_matched'],'VL Amount'),
            len(results['collection_matched'])+len(results['collection_unmatched_cl']),
            ssum(results['collection_matched'],'CL Amount')+ssum(results['collection_unmatched_cl'],'Debit'),
            len(results['collection_unmatched_vl']),
            ssum(results['collection_unmatched_vl'],'Debit')+ssum(results['collection_unmatched_vl'],'Credit'),
            pct(len(results['collection_matched']), col_vl_t), 'UTR → Period+Amount', 'FFF8EE')

    for i, w in enumerate([30, 10, 14, 10, 14, 10, 14, 12, 14, 8, 35], 1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    # Detail sheets
    write_sheet(wb, 'Inv - Matched', results['invoice_matched'], MTH_COLOR)
    write_sheet(wb, f'Inv - Unmatched VL', results['invoice_unmatched_vl'], UNM_COLOR)
    write_sheet(wb, f'Inv - Unmatched CL', results['invoice_unmatched_cl'], UNM_COLOR)
    write_sheet(wb, 'DN - Matched', dn_list, MTH_COLOR)
    write_sheet(wb, 'DN - Unmatched VL', results['dn_unmatched_vl'], UNM_COLOR)
    write_sheet(wb, 'DN - Unmatched CL', results['dn_unmatched_cl'], UNM_COLOR)
    write_sheet(wb, 'Credit Notes Matched', cn_list, MTH_COLOR)
    write_sheet(wb, 'Collections - Matched', results['collection_matched'], MTH_COLOR)
    write_sheet(wb, 'Collections - Unmatch VL', results['collection_unmatched_vl'], UNM_COLOR)
    write_sheet(wb, 'Collections - Unmatch CL', results['collection_unmatched_cl'], UNM_COLOR)
    write_sheet(wb, 'Reversals - Cross Ledger', results['reversal_cross_ledger'], '8B5000')
    write_sheet(wb, 'Reversals - VL Internal', results['reversal_vl_internal'], MTH_COLOR)
    write_sheet(wb, 'Reversals - Unmatched', results['reversal_unmatched'], UNM_COLOR)

    # VL annotated
    vl_ann = vl_orig
    ws_vl = wb.create_sheet(f'{VL[:15]} - VL'[:31])
    ws_vl.sheet_view.showGridLines = False
    vl_display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'particulars', 'debit', 'credit', 'closing'] if c in vl_ann.columns]
    vl_display_cols += ['_remark', '_match_ref']
    vl_hmap = {'doc_date': 'Doc Date', 'doc_no': 'Doc No', 'doc_type': 'Doc Type',
               'particulars': 'Particulars', 'debit': 'Debit', 'credit': 'Credit',
               'closing': 'Closing Balance', '_remark': 'Remark', '_match_ref': 'Matched With'}
    style_header(ws_vl, [vl_hmap.get(c, c) for c in vl_display_cols], row=1, color=VL_COLOR)
    for ri, (_, row) in enumerate(vl_ann[vl_display_cols].iterrows(), 2):
        remark = str(row.get('_remark', ''))
        fill = UNM_FILL if 'Unmatched' in remark else (REV_FILL if 'Reversal' in remark or 'Reversed' in remark else MTH_FILL)
        for ci, col in enumerate(vl_display_cols, 1):
            cell = write_cell(ws_vl, ri, ci, row[col])
            cell.fill = mk_fill(fill); cell.border = border
            cell.font = Font(name='Calibri', size=9); cell.alignment = Alignment(vertical='center')
        ws_vl.row_dimensions[ri].height = 18
    auto_width(ws_vl)

    # CL annotated
    cl_ann = cl_orig
    ws_cl = wb.create_sheet(f'{CL[:15]} - CL'[:31])
    ws_cl.sheet_view.showGridLines = False
    cl_display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit'] if c in cl_ann.columns]
    cl_display_cols += ['_remark', '_match_ref']
    cl_hmap = {'doc_date': 'Doc Date', 'doc_no': 'Doc No', 'doc_type': 'Doc Type',
               'debit': 'Debit', 'credit': 'Credit', '_remark': 'Remark', '_match_ref': 'Matched With'}
    style_header(ws_cl, [cl_hmap.get(c, c) for c in cl_display_cols], row=1, color=CL_COLOR)
    for ri, (_, row) in enumerate(cl_ann[cl_display_cols].iterrows(), 2):
        remark = str(row.get('_remark', ''))
        fill = UNM_FILL if 'Unmatched' in remark else MTH_FILL
        for ci, col in enumerate(cl_display_cols, 1):
            cell = write_cell(ws_cl, ri, ci, row[col])
            cell.fill = mk_fill(fill); cell.border = border
            cell.font = Font(name='Calibri', size=9); cell.alignment = Alignment(vertical='center')
        ws_cl.row_dimensions[ri].height = 18
    auto_width(ws_cl)

    wb.save(output)
    return output.getvalue()


# ─────────────────────────────────────────────
# COLUMN MAPPING REVIEW UI
# ─────────────────────────────────────────────

def render_column_mapping_ui(ledger_label, raw_columns, detected_mapping, confidence, key_prefix, tag_class):
    """
    Renders a column mapping review panel. Returns the final mapping dict {field: col}.
    """
    FIELDS = {
        'doc_date':    ('📅 Doc Date',    'Date of transaction'),
        'doc_no':      ('🔢 Doc No',      'Document / Invoice number'),
        'doc_type':    ('📋 Doc Type',    'Transaction type (Invoice, Payment, DN...)'),
        'particulars': ('📝 Particulars', 'Description / narration'),
        'debit':       ('➕ Debit',       'Debit amount column'),
        'credit':      ('➖ Credit',      'Credit amount column'),
        'closing':     ('📊 Closing Bal', 'Closing / running balance'),
    }

    OPTIONS = ['(not found)'] + raw_columns
    final_mapping = {}

    st.markdown(f'<span class="section-tag {tag_class}">{ledger_label}</span>', unsafe_allow_html=True)

    for field, (label, desc) in FIELDS.items():
        detected = detected_mapping.get(field)
        conf = confidence.get(field, 'medium') if detected else 'low'
        conf_color = 'col-confidence-high' if conf == 'high' else 'col-confidence-low'
        conf_label = f"✅ {conf.upper()}" if conf == 'high' else (f"⚠️ {conf.upper()}" if conf == 'medium' else "❌ LOW")

        default_idx = OPTIONS.index(detected) if detected in OPTIONS else 0

        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            st.markdown(f"**{label}**  \n<small style='color:#888'>{desc}</small>", unsafe_allow_html=True)
        with col2:
            selected = st.selectbox(
                label,
                OPTIONS,
                index=default_idx,
                key=f"{key_prefix}_{field}",
                label_visibility='collapsed',
            )
        with col3:
            if detected:
                st.markdown(f"<span class='{conf_color}' style='font-size:0.7rem'>{conf_label}</span>", unsafe_allow_html=True)
            else:
                st.markdown("<span style='color:#A32035;font-size:0.7rem'>❌ NOT FOUND</span>", unsafe_allow_html=True)

        if selected != '(not found)':
            final_mapping[field] = selected

    return final_mapping


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────

def main():
    if 'vname' not in st.session_state:
        st.session_state['vname'] = ''
    if 'cname' not in st.session_state:
        st.session_state['cname'] = ''

    st.markdown("""
    <div class="recon-header">
        <div>
            <div class="recon-logo">⚖️ Ledger Reconciliation</div>
            <div class="recon-subtitle">Vendor · Customer Ledger Reconciliation · For Indian CAs &amp; CFOs &nbsp;&nbsp; <span class="ai-badge">✨ AI Column Detection</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    nc1, nc2, _ = st.columns([2, 2, 1])
    with nc1:
        vname = st.text_input("🏭 Vendor Name", value=st.session_state['vname'], placeholder="e.g. ABC Suppliers Pvt. Ltd.")
        if vname:
            st.session_state['vname'] = vname
    with nc2:
        cname = st.text_input("🏢 Customer Name", value=st.session_state['cname'], placeholder="e.g. XYZ Traders Ltd.")
        if cname:
            st.session_state['cname'] = cname

    vname = st.session_state.get('vname', 'Vendor') or 'Vendor'
    cname = st.session_state.get('cname', 'Customer') or 'Customer'
    VL = vname
    CL = cname

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        st.markdown("---")
        tolerance = st.number_input(
            "Amount Tolerance (₹)",
            min_value=0.0, max_value=100.0, value=1.0, step=0.5,
            help="Max difference allowed when matching by amount"
        )
        st.markdown("---")
        st.markdown("### 📋 Matching Rules")
        st.markdown(f"""
        <div class="info-box">
        1. Invoices → Doc Number<br>
        2. Reversals → Complete Reversal<br>
        3. Credit Notes ({VL}) → Discount DN / PRN ({CL})<br>
        4. Debit Notes → Doc No → Period+Amount<br>
        5. Collections → UTR → Period+Amount<br>
        6. Remaining → Unmatched
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("""
        <div class="info-box" style="font-size:0.72rem;">
        <b>✨ AI Column Detection</b><br>
        Columns are auto-detected using Claude AI. Any format from Tally, SAP, Oracle, Zoho, or custom ERP is supported.<br><br>
        After upload, review the detected column mapping before running reconciliation.
        </div>
        """, unsafe_allow_html=True)

    # File upload
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<span class="section-tag tag-vl">📘 {VL.upper()} — VENDOR LEDGER</span>', unsafe_allow_html=True)
        vl_file = st.file_uploader(f"Upload {VL} Ledger (.xlsx / .xls)", type=['xlsx', 'xls'], key='vl')
    with col2:
        st.markdown(f'<span class="section-tag tag-cl">📗 {CL.upper()} — CUSTOMER LEDGER</span>', unsafe_allow_html=True)
        cl_file = st.file_uploader(f"Upload {CL} Ledger (.xlsx / .xls)", type=['xlsx', 'xls'], key='cl')

    if not vl_file or not cl_file:
        st.markdown(f"""
        <div class="info-box" style="margin-top:2rem; border-left-color: #4f8eff;">
        📂 Upload both ledger files above to begin.<br>
        <b>✨ AI-powered column detection</b> will automatically identify your columns — works with any ERP format.
        </div>
        """, unsafe_allow_html=True)
        for key in ['results', 'excel_data', 'vl_parsed', 'cl_parsed', 'vl_raw_cols', 'cl_raw_cols',
                    'vl_detected_map', 'cl_detected_map', 'vl_confidence', 'cl_confidence']:
            st.session_state.pop(key, None)
        return

    # Parse files (cache by file identity)
    vl_file_id = (vl_file.name, vl_file.size)
    cl_file_id = (cl_file.name, cl_file.size)

    if st.session_state.get('vl_file_id') != vl_file_id or 'vl_parsed' not in st.session_state:
        with st.spinner(f"🤖 Analysing {VL} ledger columns with AI..."):
            try:
                df, closing, mapping, raw_cols, confidence = load_vendor_ledger(vl_file)
                st.session_state['vl_parsed'] = df
                st.session_state['vl_closing'] = closing
                st.session_state['vl_detected_map'] = mapping
                st.session_state['vl_raw_cols'] = raw_cols
                st.session_state['vl_confidence'] = confidence
                st.session_state['vl_file_id'] = vl_file_id
                st.session_state.pop('results', None)
                st.session_state.pop('excel_data', None)
            except Exception as e:
                st.error(f"Error reading {VL} file: {e}")
                return

    if st.session_state.get('cl_file_id') != cl_file_id or 'cl_parsed' not in st.session_state:
        with st.spinner(f"🤖 Analysing {CL} ledger columns with AI..."):
            try:
                df, closing, mapping, raw_cols, confidence = load_customer_ledger(cl_file)
                st.session_state['cl_parsed'] = df
                st.session_state['cl_closing'] = closing
                st.session_state['cl_detected_map'] = mapping
                st.session_state['cl_raw_cols'] = raw_cols
                st.session_state['cl_confidence'] = confidence
                st.session_state['cl_file_id'] = cl_file_id
                st.session_state.pop('results', None)
                st.session_state.pop('excel_data', None)
            except Exception as e:
                st.error(f"Error reading {CL} file: {e}")
                return

    vl = st.session_state['vl_parsed']
    cl = st.session_state['cl_parsed']
    vl_closing = st.session_state.get('vl_closing')
    cl_closing = st.session_state.get('cl_closing')

    st.success(f"✅ Files parsed — {VL}: **{len(vl)} rows** · {CL}: **{len(cl)} rows**")

    cb1, cb2 = st.columns(2)
    with cb1:
        st.info(f"📘 **{VL} Closing Balance:** {fmt_inr(vl_closing) if vl_closing is not None else 'Not detected'}")
    with cb2:
        st.info(f"📗 **{CL} Closing Balance:** {fmt_inr(cl_closing) if cl_closing is not None else 'Not detected'}")

    # ── COLUMN MAPPING REVIEW (KEY FEATURE) ──
    with st.expander("🤖 Review AI-Detected Column Mapping", expanded=True):
        st.markdown("""
        <div class="info-box">
        <b>✨ AI has automatically detected your column mapping.</b> Please review and adjust if needed before running reconciliation.
        Incorrect column mapping is the #1 cause of low match rates — verify especially <b>Doc No</b>, <b>Debit</b>, and <b>Credit</b> columns.
        </div>
        """, unsafe_allow_html=True)

        map_col1, map_col2 = st.columns(2)

        with map_col1:
            vl_final_map = render_column_mapping_ui(
                f"📘 {VL} — Vendor Ledger",
                st.session_state['vl_raw_cols'],
                st.session_state['vl_detected_map'],
                st.session_state['vl_confidence'],
                key_prefix='vl',
                tag_class='tag-vl',
            )

        with map_col2:
            cl_final_map = render_column_mapping_ui(
                f"📗 {CL} — Customer Ledger",
                st.session_state['cl_raw_cols'],
                st.session_state['cl_detected_map'],
                st.session_state['cl_confidence'],
                key_prefix='cl',
                tag_class='tag-cl',
            )

        if st.button("✅ Apply Column Mapping & Re-parse", key='apply_col_map'):
            with st.spinner("Re-parsing with updated column mapping..."):
                try:
                    vl_file.seek(0)
                    df_vl, closing_vl, _, raw_cols_vl, conf_vl = _load_any_ledger_smart(vl_file, is_vendor=True, override_mapping=vl_final_map)
                    st.session_state['vl_parsed'] = df_vl
                    st.session_state['vl_closing'] = closing_vl

                    cl_file.seek(0)
                    df_cl, closing_cl, _, raw_cols_cl, conf_cl = _load_any_ledger_smart(cl_file, is_vendor=False, override_mapping=cl_final_map)
                    st.session_state['cl_parsed'] = df_cl
                    st.session_state['cl_closing'] = closing_cl

                    st.session_state.pop('results', None)
                    st.session_state.pop('excel_data', None)
                    st.success("✅ Column mapping applied. Click Run Reconciliation.")
                except Exception as e:
                    st.error(f"Error applying mapping: {e}")

    # Preview
    with st.expander("👁 Preview Parsed Data"):
        pc1, pc2 = st.columns(2)
        with pc1:
            st.markdown(f'<span class="section-tag tag-vl">{VL} — VENDOR LEDGER</span>', unsafe_allow_html=True)
            display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'particulars', 'debit', 'credit', 'closing'] if c in vl.columns]
            st.dataframe(vl[display_cols].head(10), use_container_width=True, hide_index=True)
        with pc2:
            st.markdown(f'<span class="section-tag tag-cl">{CL} — CUSTOMER LEDGER</span>', unsafe_allow_html=True)
            display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit'] if c in cl.columns]
            st.dataframe(cl[display_cols].head(10), use_container_width=True, hide_index=True)

    if st.button("▶ Run Reconciliation", use_container_width=False):
        with st.spinner("Running reconciliation engine..."):
            results = run_reconciliation(vl, cl, tolerance=tolerance)

        def safe_records(df):
            if not isinstance(df, pd.DataFrame):
                return df
            d = df.copy()
            for col in list(d.columns):
                try:
                    if col.startswith('_'):
                        d[col] = d[col].astype(str).replace('nan', '').replace('None', '')
                    elif hasattr(d[col], 'dtype') and d[col].dtype == object:
                        d[col] = d[col].fillna('').astype(str)
                    elif hasattr(d[col], 'dtype') and str(d[col].dtype).startswith('datetime'):
                        d[col] = d[col].astype(str)
                except Exception:
                    try:
                        d[col] = d[col].fillna('').astype(str)
                    except Exception:
                        pass
            return d.to_dict('records')

        results['vl_annotated'] = safe_records(results['vl_annotated'])
        results['cl_annotated'] = safe_records(results['cl_annotated'])
        results['vl_closing'] = float(vl_closing) if vl_closing is not None else None
        results['cl_closing'] = float(cl_closing) if cl_closing is not None else None
        results['vl_name'] = VL
        results['cl_name'] = CL
        st.session_state['results'] = results
        st.session_state.pop('excel_data', None)
        st.session_state.pop('excel_key', None)

    if 'results' not in st.session_state:
        return

    results = st.session_state['results']
    VL = results.get('vl_name', vname) or vname
    CL = results.get('cl_name', cname) or cname
    vl_ann_df = pd.DataFrame(results['vl_annotated'])
    cl_ann_df = pd.DataFrame(results['cl_annotated'])
    vl_closing_val = results.get('vl_closing')
    cl_closing_val = results.get('cl_closing')

    cn_matched = [r for r in results['dn_matched'] if 'Credit Note' in str(r.get('Match Type', ''))]
    dn_only_matched = [r for r in results['dn_matched'] if 'Credit Note' not in str(r.get('Match Type', ''))]

    inv_matched_cnt = len(results['invoice_matched'])
    inv_un_vl_cnt   = len(results['invoice_unmatched_vl'])
    inv_un_cl_cnt   = len(results['invoice_unmatched_cl'])
    dn_matched_cnt  = len(dn_only_matched)
    col_matched_cnt = len(results['collection_matched'])
    cn_matched_cnt  = len(cn_matched)

    total_matched_cnt = inv_matched_cnt + dn_matched_cnt + cn_matched_cnt + col_matched_cnt
    total_un_vl_cnt   = inv_un_vl_cnt + len(results['dn_unmatched_vl']) + len(results['collection_unmatched_vl'])
    total_un_cl_cnt   = inv_un_cl_cnt + len(results['dn_unmatched_cl']) + len(results['collection_unmatched_cl'])

    inv_matched_val = safe_sum(results['invoice_matched'], 'VL Debit') + safe_sum(results['invoice_matched'], 'VL Credit')
    inv_un_vl_val   = safe_sum(results['invoice_unmatched_vl'], 'Debit') + safe_sum(results['invoice_unmatched_vl'], 'Credit')
    inv_un_cl_val   = safe_sum(results['invoice_unmatched_cl'], 'Debit') + safe_sum(results['invoice_unmatched_cl'], 'Credit')
    col_matched_val = safe_sum(results['collection_matched'], 'VL Amount')
    col_un_vl_val   = safe_sum(results['collection_unmatched_vl'], 'Debit') + safe_sum(results['collection_unmatched_vl'], 'Credit')
    col_un_cl_val   = safe_sum(results['collection_unmatched_cl'], 'Debit') + safe_sum(results['collection_unmatched_cl'], 'Credit')
    cn_matched_val  = safe_sum(cn_matched, 'VL Debit') + safe_sum(cn_matched, 'VL Credit')
    dn_matched_val  = safe_sum(dn_only_matched, 'VL Debit') + safe_sum(dn_only_matched, 'VL Credit')
    total_matched_val = inv_matched_val + dn_matched_val + cn_matched_val + col_matched_val
    total_un_vl_val = inv_un_vl_val + safe_sum(results['dn_unmatched_vl'], 'Debit') + col_un_vl_val
    total_un_cl_val = inv_un_cl_val + safe_sum(results['dn_unmatched_cl'], 'Debit') + col_un_cl_val

    # Stats cards
    st.markdown(f"""
    <div class="stat-grid">
        <div class="stat-card total">
            <div class="stat-label">Total Items</div>
            <div class="stat-value">{total_matched_cnt + total_un_vl_cnt}</div>
            <div class="stat-sub">VL items processed</div>
        </div>
        <div class="stat-card matched">
            <div class="stat-label">Matched</div>
            <div class="stat-value">{total_matched_cnt}</div>
            <div class="stat-sub">{fmt_inr(total_matched_val)}</div>
        </div>
        <div class="stat-card unmatched">
            <div class="stat-label">Unmatched VL</div>
            <div class="stat-value">{total_un_vl_cnt}</div>
            <div class="stat-sub">{fmt_inr(total_un_vl_val)}</div>
        </div>
        <div class="stat-card partial">
            <div class="stat-label">Unmatched CL</div>
            <div class="stat-value">{total_un_cl_cnt}</div>
            <div class="stat-sub">{fmt_inr(total_un_cl_val)}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Excel report
    if 'excel_data' not in st.session_state or st.session_state.get('excel_key') != id(results):
        try:
            st.session_state['excel_data'] = build_excel(results, vl_ann_df, cl_ann_df, VL, CL)
            st.session_state['excel_key'] = id(results)
        except Exception as e:
            st.session_state['excel_data'] = None
            st.error(f"Error generating Excel: {e}")

    excel_data = st.session_state.get('excel_data')
    if excel_data:
        st.download_button(
            label="⬇️  Download Reconciliation Report (.xlsx)",
            data=excel_data,
            file_name=f"Recon_{VL}_{CL}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_top",
        )

    # Results tabs
    tabs = st.tabs([
        "🧾 Invoices",
        "📝 DN / Credit Notes",
        "💰 Collections",
        "🔁 Reversals",
        "⚠️ All Unmatched",
        f"📘 {VL} Ledger",
        f"📗 {CL} Ledger",
    ])

    with tabs[0]:
        st.markdown(f'<span class="section-tag tag-matched">MATCHED INVOICES — {VL} vs {CL}</span>', unsafe_allow_html=True)
        display_df(results['invoice_matched'])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f'<span class="section-tag tag-vl">UNMATCHED — {VL}</span>', unsafe_allow_html=True)
            display_df(results['invoice_unmatched_vl'])
        with c2:
            st.markdown(f'<span class="section-tag tag-cl">UNMATCHED — {CL}</span>', unsafe_allow_html=True)
            display_df(results['invoice_unmatched_cl'])

    with tabs[1]:
        st.markdown(f'<span class="section-tag tag-blue">CREDIT NOTES ({VL}) ↔ DISCOUNT DN / PRN ({CL})</span>', unsafe_allow_html=True)
        display_df(cn_matched)
        st.markdown("---")
        st.markdown(f'<span class="section-tag tag-matched">MATCHED DEBIT NOTES</span>', unsafe_allow_html=True)
        display_df(dn_only_matched)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f'<span class="section-tag tag-vl">UNMATCHED DN — {VL}</span>', unsafe_allow_html=True)
            display_df(results['dn_unmatched_vl'])
        with c2:
            st.markdown(f'<span class="section-tag tag-cl">UNMATCHED DN — {CL}</span>', unsafe_allow_html=True)
            display_df(results['dn_unmatched_cl'])

    with tabs[2]:
        st.markdown(f'<span class="section-tag tag-matched">MATCHED COLLECTIONS</span>', unsafe_allow_html=True)
        display_df(results['collection_matched'])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f'<span class="section-tag tag-vl">UNMATCHED — {VL}</span>', unsafe_allow_html=True)
            display_df(results['collection_unmatched_vl'])
        with c2:
            st.markdown(f'<span class="section-tag tag-cl">UNMATCHED — {CL}</span>', unsafe_allow_html=True)
            display_df(results['collection_unmatched_cl'])

    with tabs[3]:
        st.markdown(f'<span class="section-tag tag-partial">REVERSED IN {VL} | INVOICE ALSO IN {CL}</span>', unsafe_allow_html=True)
        display_df(results['reversal_cross_ledger'])
        st.markdown("---")
        st.markdown(f'<span class="section-tag tag-matched">REVERSED IN {VL} | NOT IN {CL}</span>', unsafe_allow_html=True)
        display_df(results['reversal_vl_internal'])
        st.markdown("---")
        rev_mis  = [r for r in results['reversal_unmatched'] if r.get('Reason', '') == 'Amount Mismatch']
        rev_miss = [r for r in results['reversal_unmatched'] if r.get('Reason', '') != 'Amount Mismatch']
        st.markdown(f'<span class="section-tag tag-unmatched">REVERSAL — AMOUNT MISMATCH ({len(rev_mis)})</span>', unsafe_allow_html=True)
        display_df(rev_mis)
        st.markdown(f'<span class="section-tag tag-unmatched">REVERSAL — ORIGINAL NOT FOUND ({len(rev_miss)})</span>', unsafe_allow_html=True)
        display_df(rev_miss)

    with tabs[4]:
        all_unmatched = []
        for item in results['invoice_unmatched_vl'] + results['dn_unmatched_vl'] + results['collection_unmatched_vl']:
            item = dict(item); item['Ledger'] = VL; all_unmatched.append(item)
        for item in results['invoice_unmatched_cl'] + results['dn_unmatched_cl'] + results['collection_unmatched_cl']:
            item = dict(item); item['Ledger'] = CL; all_unmatched.append(item)
        st.markdown(f'<span class="section-tag tag-unmatched">ALL UNMATCHED — {VL} & {CL}</span>', unsafe_allow_html=True)
        display_df(all_unmatched)

    with tabs[5]:
        st.markdown(f'<span class="section-tag tag-vl">{VL} — VENDOR LEDGER WITH REMARKS</span>', unsafe_allow_html=True)
        st.caption("🟢 Green = Matched | 🔴 Red = Unmatched | 🟡 Yellow = Reversal")
        if not vl_ann_df.empty:
            disp = vl_ann_df.copy()
            for col in disp.columns:
                if 'date' in col.lower():
                    disp[col] = pd.to_datetime(disp[col], errors='coerce').dt.strftime('%d-%b-%Y').fillna('')
            st.dataframe(disp, use_container_width=True, hide_index=True)
            if vl_closing_val:
                st.info(f"**{VL} Closing Balance: {fmt_inr(vl_closing_val)}**")

    with tabs[6]:
        st.markdown(f'<span class="section-tag tag-cl">{CL} — CUSTOMER LEDGER WITH REMARKS</span>', unsafe_allow_html=True)
        st.caption("🟢 Green = Matched | 🔴 Red = Unmatched")
        if not cl_ann_df.empty:
            disp = cl_ann_df.copy()
            for col in disp.columns:
                if 'date' in col.lower():
                    disp[col] = pd.to_datetime(disp[col], errors='coerce').dt.strftime('%d-%b-%Y').fillna('')
            st.dataframe(disp, use_container_width=True, hide_index=True)

    # Reconciliation Statement
    st.markdown("---")
    st.markdown("### 📋 Ledger Reconciliation Statement")
    vl_bal     = float(vl_closing_val) if vl_closing_val else 0.0
    cl_bal_act = float(cl_closing_val) if cl_closing_val is not None else 0.0
    adj_inv_vl = inv_un_vl_val
    adj_cn_vl  = cn_matched_val
    adj_dn_cl  = safe_sum(results['dn_unmatched_cl'], 'Debit') + safe_sum(results['dn_unmatched_cl'], 'Credit')
    adj_inv_cl = inv_un_cl_val
    adj_pay_vl = col_un_vl_val
    adj_pay_cl = col_un_cl_val
    net_bal_b  = vl_bal - adj_inv_vl + adj_cn_vl - adj_dn_cl + adj_inv_cl + adj_pay_vl - adj_pay_cl
    diff_bc    = net_bal_b - cl_bal_act

    recon_rows = [
        ("Particular", "Amount", "col_header"),
        (f"Balance as per {VL} Books (A)", fmt_inr(vl_bal), "header"),
        (f"Less: Tax invoice in {VL} but not in {CL}", fmt_inr(-adj_inv_vl), "less"),
        (f"Add: Credit note in {VL} but not in {CL}", fmt_inr(adj_cn_vl), "add"),
        (f"Less: Debit notes in {CL} but not in {VL}", fmt_inr(-adj_dn_cl), "less"),
        (f"Add: Tax invoice in {CL} but not in {VL}", fmt_inr(adj_inv_cl), "add"),
        (f"Add: Payment not available in {CL}", fmt_inr(adj_pay_vl), "add"),
        (f"Less: Payment not available in {VL}", fmt_inr(-adj_pay_cl), "less"),
        ("", "", "blank"),
        (f"Net Balance as per {VL} Books — B", fmt_inr(net_bal_b), "total"),
        ("", "", "blank"),
        (f"Balance as per {CL} Books — C", fmt_inr(cl_bal_act), "cl_total"),
        ("", "", "blank"),
        ("Unreconciled Difference B - C", fmt_inr(diff_bc), "diff"),
        ("(Should be zero after all adjustments)", "", "note"),
    ]
    styles = {
        "col_header": ("#1c2130", "#4f8eff", True),
        "header":     ("#1a3a6b", "#ffffff", True),
        "less":       ("#1a1a2a", "#ff9999", False),
        "add":        ("#1a2a1a", "#99ffcc", False),
        "blank":      ("#0d0f14", "#0d0f14", False),
        "total":      ("#1a3a6b", "#ffffff", True),
        "cl_total":   ("#1a5a3a", "#ffffff", True),
        "diff":       ("#cc0000", "#ffffff", True),
        "note":       ("#cc0000", "#ffcccc", False),
    }
    rs1, rs2 = st.columns([3, 1])
    for label, amount, row_type in recon_rows:
        bg, fg, bold = styles.get(row_type, ("#141720", "#e8ecf4", False))
        fw = "700" if bold else "400"
        with rs1:
            st.markdown(f"<div style='background:{bg};color:{fg};padding:7px 14px;border-bottom:1px solid #252c3d;font-size:0.83rem;font-weight:{fw}'>{label}</div>", unsafe_allow_html=True)
        with rs2:
            st.markdown(f"<div style='background:{bg};color:{fg};padding:7px 14px;border-bottom:1px solid #252c3d;font-size:0.83rem;font-weight:{fw};text-align:right'>{amount}</div>", unsafe_allow_html=True)

    st.markdown("---")
    if excel_data:
        st.download_button(
            label="⬇️  Download Reconciliation Report (.xlsx)",
            data=excel_data,
            file_name=f"Recon_{VL}_{CL}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
            key="download_bottom",
        )


if __name__ == "__main__":
    main()
