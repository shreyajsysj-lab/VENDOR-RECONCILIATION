import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="VendorSync · Reconciliation",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

:root {
    --bg: #0d0f14;
    --surface: #141720;
    --surface2: #1c2130;
    --border: #252c3d;
    --accent: #4f8eff;
    --accent2: #00d4aa;
    --warn: #ff8c42;
    --danger: #ff4d6d;
    --text: #e8ecf4;
    --muted: #7a8499;
    --matched: #00d4aa22;
    --matched-border: #00d4aa55;
    --unmatched: #ff4d6d22;
    --unmatched-border: #ff4d6d55;
    --partial: #ff8c4222;
    --partial-border: #ff8c4255;
}

html, body, [class*="css"] {
    font-family: 'DM Mono', monospace;
    background-color: var(--bg) !important;
    color: var(--text) !important;
}

.main { background: var(--bg) !important; }
.block-container { padding: 1.5rem 2rem !important; max-width: 1400px; }

[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] * { color: var(--text) !important; }

.recon-header {
    display: flex;
    align-items: center;
    gap: 1rem;
    margin-bottom: 2rem;
    padding-bottom: 1.5rem;
    border-bottom: 1px solid var(--border);
}
.recon-logo {
    font-family: 'Syne', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: -0.03em;
}
.recon-subtitle {
    font-size: 0.75rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.12em;
}

.stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.5rem; }
.stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1rem 1.25rem;
    position: relative;
    overflow: hidden;
}
.stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
}
.stat-card.matched::before { background: var(--accent2); }
.stat-card.unmatched::before { background: var(--danger); }
.stat-card.partial::before { background: var(--warn); }
.stat-card.total::before { background: var(--accent); }
.stat-label { font-size: 0.7rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; }
.stat-value { font-family: 'Syne', sans-serif; font-size: 1.8rem; font-weight: 700; margin: 0.25rem 0; }
.stat-card.matched .stat-value { color: var(--accent2); }
.stat-card.unmatched .stat-value { color: var(--danger); }
.stat-card.partial .stat-value { color: var(--warn); }
.stat-card.total .stat-value { color: var(--accent); }
.stat-sub { font-size: 0.7rem; color: var(--muted); }

/* Summary table */
.summary-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-bottom: 1.5rem;
}
.summary-table th {
    background: var(--surface2);
    color: var(--accent);
    padding: 10px 14px;
    text-align: left;
    border-bottom: 2px solid var(--border);
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 0.7rem;
}
.summary-table td {
    padding: 9px 14px;
    border-bottom: 1px solid var(--border);
    color: var(--text);
}
.summary-table tr:hover td { background: var(--surface2); }
.summary-table .num { text-align: right; font-weight: 600; }
.summary-table .matched-val { color: var(--accent2); }
.summary-table .unmatched-val { color: var(--danger); }
.summary-table .total-row td { background: var(--surface2); font-weight: 700; border-top: 2px solid var(--border); }

.stTabs [data-baseweb="tab-list"] {
    background: var(--surface) !important;
    border-radius: 8px;
    padding: 4px;
    gap: 2px;
    border: 1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Syne', sans-serif !important;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    color: var(--muted) !important;
    background: transparent !important;
    border-radius: 6px !important;
    padding: 6px 16px !important;
}
.stTabs [aria-selected="true"] {
    background: var(--surface2) !important;
    color: var(--text) !important;
}

[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

.stButton > button {
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    background: linear-gradient(135deg, var(--accent), #3d6fcc) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.2s !important;
}
.stButton > button:hover { opacity: 0.88 !important; transform: translateY(-1px); }

.stDownloadButton > button {
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    background: var(--surface2) !important;
    color: var(--accent2) !important;
    border: 1px solid var(--accent2) !important;
    border-radius: 8px !important;
}

[data-testid="stFileUploader"] {
    background: var(--surface) !important;
    border: 1px dashed var(--border) !important;
    border-radius: 10px !important;
}

.section-tag {
    display: inline-block;
    font-family: 'Syne', sans-serif;
    font-size: 0.65rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.15em;
    padding: 3px 10px;
    border-radius: 4px;
    margin-bottom: 0.5rem;
}
.tag-matched { background: var(--matched); color: var(--accent2); border: 1px solid var(--matched-border); }
.tag-unmatched { background: var(--unmatched); color: var(--danger); border: 1px solid var(--unmatched-border); }
.tag-partial { background: var(--partial); color: var(--warn); border: 1px solid var(--partial-border); }
.tag-blue { background: #4f8eff22; color: var(--accent); border: 1px solid #4f8eff55; }

.info-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-left: 3px solid var(--accent);
    border-radius: 6px;
    padding: 0.75rem 1rem;
    font-size: 0.8rem;
    color: var(--muted);
    margin-bottom: 1rem;
}

[data-testid="stAlert"] { border-radius: 8px !important; }

[data-testid="stSelectbox"] > div, [data-testid="stMultiSelect"] > div {
    background: var(--surface) !important;
    border-color: var(--border) !important;
    border-radius: 8px !important;
}

[data-testid="stNumberInput"] input {
    background: var(--surface) !important;
    border-color: var(--border) !important;
    color: var(--text) !important;
    border-radius: 8px !important;
}

[data-testid="stExpander"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────

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
    return any(k in s for k in ['DEBIT NOTE', 'DN', 'DEBIT MEMO', 'DM', 'CREDIT MEMO', 'CM'])

def is_reversal_type(doc_type):
    """Detect Complete Reversal or Saleable Return doc types."""
    if pd.isna(doc_type):
        return False
    s = str(doc_type).upper()
    return any(k in s for k in [
        'COMPLETE REVERSAL', 'REVERSAL', 'SALEABLE RETURN',
        'SALES RETURN', 'RETURN', 'CANCELLATION', 'CANCEL', 'REVERSED'
    ])

def is_collection(doc_type, particulars=""):
    if pd.isna(doc_type):
        doc_type = ""
    s = str(doc_type).upper() + " " + str(particulars).upper()
    return any(k in s for k in ['PAYMENT', 'RECEIPT', 'COLLECTION', 'NEFT', 'RTGS', 'IMPS', 'CHEQUE', 'CHQ', 'TDS', 'BANK', 'UTR'])

def extract_ref_from_particulars(particulars):
    """Extract referenced invoice number from particulars column."""
    if pd.isna(particulars):
        return ""
    s = str(particulars).strip().upper()
    # Remove common prefix words
    s = re.sub(r'(REVERSAL OF|AGAINST|REF|REFERENCE|RETURN OF|CANCELLATION OF|REVERSED)\s*', '', s)
    # Clean up
    s = re.sub(r'[\s\-_/]', '', s)
    return s.strip()

# ─────────────────────────────────────────────
# LOAD & PARSE
# ─────────────────────────────────────────────

def load_vendor_ledger(file):
    df = pd.read_excel(file, header=None)
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).lower() for v in row if not pd.isna(v)]
        if any('doc' in v and 'date' in v for v in vals) or any('doc. date' in v for v in vals):
            header_row = i
            break
    if header_row is None:
        for i, row in df.iterrows():
            vals = [str(v).lower() for v in row if not pd.isna(v)]
            if any('debit' in v for v in vals) and any('credit' in v for v in vals):
                header_row = i
                break
    if header_row is None:
        header_row = 0

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'doc' in cl and 'date' in cl:
            col_map[c] = 'doc_date'
        elif 'doc' in cl and ('no' in cl or 'num' in cl or 'type' not in cl) and 'date' not in cl and 'type' not in cl:
            col_map[c] = 'doc_no'
        elif 'doc' in cl and 'type' in cl:
            col_map[c] = 'doc_type'
        elif 'particular' in cl:
            col_map[c] = 'particulars'
        elif 'opening' in cl:
            col_map[c] = 'opening'
        elif 'debit' in cl:
            col_map[c] = 'debit'
        elif 'credit' in cl:
            col_map[c] = 'credit'
        elif 'closing' in cl or 'balance' in cl:
            col_map[c] = 'closing'
    df = df.rename(columns=col_map)

    needed = ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit', 'particulars']
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    df['doc_date'] = pd.to_datetime(df['doc_date'], errors='coerce', dayfirst=True)
    df['debit'] = pd.to_numeric(df['debit'], errors='coerce').fillna(0)
    df['credit'] = pd.to_numeric(df['credit'], errors='coerce').fillna(0)
    df['doc_no_clean'] = df['doc_no'].apply(clean_doc_number)
    df['period'] = df['doc_date'].apply(get_period)
    df['particulars_ref'] = df['particulars'].apply(extract_ref_from_particulars)
    df = df[df['doc_no_clean'] != ''].reset_index(drop=True)
    df['_idx'] = df.index
    df['_remark'] = ''
    df['_match_ref'] = ''
    return df


def load_customer_ledger(file):
    df = pd.read_excel(file, header=None)
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).lower() for v in row if not pd.isna(v)]
        if any('document' in v for v in vals) and any('debit' in v or 'credit' in v for v in vals):
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'date' in cl:
            col_map[c] = 'doc_date'
        elif 'type' in cl:
            col_map[c] = 'doc_type'
        elif ('no' in cl or 'detail' in cl or 'num' in cl) and 'date' not in cl and 'type' not in cl:
            col_map[c] = 'doc_no'
        elif 'debit' in cl:
            col_map[c] = 'debit'
        elif 'credit' in cl:
            col_map[c] = 'credit'
    df = df.rename(columns=col_map)

    needed = ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit']
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    df['doc_date'] = pd.to_datetime(df['doc_date'], errors='coerce', dayfirst=True)
    df['debit'] = pd.to_numeric(df['debit'], errors='coerce').fillna(0)
    df['credit'] = pd.to_numeric(df['credit'], errors='coerce').fillna(0)
    df['doc_no_clean'] = df['doc_no'].apply(clean_doc_number)
    df['period'] = df['doc_date'].apply(get_period)
    df = df[df['doc_no_clean'] != ''].reset_index(drop=True)
    df['_idx'] = df.index
    df['_remark'] = ''
    df['_match_ref'] = ''
    return df

# ─────────────────────────────────────────────
# RECONCILIATION ENGINE
# ─────────────────────────────────────────────

def run_reconciliation(vl_orig, cl_orig, tolerance=1.0):
    results = {
        'invoice_matched': [],
        'invoice_unmatched_vl': [],
        'invoice_unmatched_cl': [],
        'dn_matched': [],
        'dn_unmatched_vl': [],
        'dn_unmatched_cl': [],
        'collection_matched': [],
        'collection_unmatched_vl': [],
        'collection_unmatched_cl': [],
        'reversal_matched': [],
        'reversal_unmatched': [],
    }

    vl = vl_orig.copy()
    cl = cl_orig.copy()
    vl['_matched'] = False
    cl['_matched'] = False

    # ── STEP 1: Handle Reversal entries in Vendor Ledger ──
    # These are rows with doc_type = Complete Reversal / Saleable Return
    # They reference original invoice in the Particulars column
    vl_reversals = vl[vl['doc_type'].apply(is_reversal_type)].copy()

    for idx, rev_row in vl_reversals.iterrows():
        ref = rev_row.get('particulars_ref', '')
        doc_no = rev_row.get('doc_no_clean', '')

        # Find the original invoice in VL being reversed (by particulars ref or doc_no)
        orig_candidates = vl[
            (~vl['_matched']) &
            (~vl['doc_type'].apply(is_reversal_type)) &
            (~vl['doc_type'].apply(is_debit_note)) &
            (~vl['doc_type'].apply(lambda x: is_collection(x))) &
            ((vl['doc_no_clean'] == ref) | (vl['doc_no_clean'] == doc_no))
        ]

        if not orig_candidates.empty:
            orig_row = orig_candidates.iloc[0]
            # Mark both the original and reversal as matched (reversed pair)
            vl.at[orig_row['_idx'], '_matched'] = True
            vl.at[orig_row['_idx'], '_remark'] = 'Matched - Reversed'
            vl.at[orig_row['_idx'], '_match_ref'] = str(rev_row.get('doc_no', ''))
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = 'Matched - Reversal Entry'
            vl.at[idx, '_match_ref'] = str(orig_row.get('doc_no', ''))

            results['reversal_matched'].append({
                'Original Doc No': orig_row.get('doc_no', ''),
                'Original Date': orig_row.get('doc_date', ''),
                'Original Type': orig_row.get('doc_type', ''),
                'Original Debit': orig_row.get('debit', 0),
                'Original Credit': orig_row.get('credit', 0),
                'Reversal Doc No': rev_row.get('doc_no', ''),
                'Reversal Date': rev_row.get('doc_date', ''),
                'Reversal Type': rev_row.get('doc_type', ''),
                'Reversal Debit': rev_row.get('debit', 0),
                'Reversal Credit': rev_row.get('credit', 0),
                'Particulars': rev_row.get('particulars', ''),
                'Remark': 'Reversed & Reconciled',
            })
        else:
            # Reversal with no matching original — mark as unmatched
            vl.at[idx, '_matched'] = True  # exclude from further matching
            vl.at[idx, '_remark'] = 'Unmatched - Reversal (No Original Found)'
            results['reversal_unmatched'].append({
                'Doc No': rev_row.get('doc_no', ''),
                'Date': rev_row.get('doc_date', ''),
                'Type': rev_row.get('doc_type', ''),
                'Particulars': rev_row.get('particulars', ''),
                'Debit': rev_row.get('debit', 0),
                'Credit': rev_row.get('credit', 0),
                'Remark': 'Reversal Entry - Original Not Found',
            })

    # ── STEP 2: Match Invoices by Document Number ──
    vl_inv = vl[
        (~vl['_matched']) &
        (~vl['doc_type'].apply(is_debit_note)) &
        (~vl['doc_type'].apply(lambda x: is_collection(x, ''))) &
        (~vl['doc_type'].apply(is_reversal_type))
    ].copy()

    cl_inv = cl[
        (~cl['_matched']) &
        (~cl['doc_type'].apply(is_debit_note)) &
        (~cl['doc_type'].apply(lambda x: is_collection(x, '')))
    ].copy()

    for idx, vrow in vl_inv.iterrows():
        matches = cl_inv[(cl_inv['doc_no_clean'] == vrow['doc_no_clean']) & (~cl_inv['_matched'])]
        if not matches.empty:
            crow = matches.iloc[0]
            vl.at[idx, '_matched'] = True
            vl.at[idx, '_remark'] = 'Matched'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = 'Matched'
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
                'Remark': 'Matched',
            })

    # ── STEP 3: Match Debit Notes ──
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
            vl.at[idx, '_remark'] = 'Matched'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = 'Matched'
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
                'Remark': 'Matched',
            })

    # ── STEP 4: Match Collections ──
    vl_col = vl[(~vl['_matched']) & (vl['doc_type'].apply(lambda x: is_collection(x, '')))].copy()
    cl_col = cl[(~cl['_matched']) & (cl['doc_type'].apply(lambda x: is_collection(x, '')))].copy()

    vl_col['utr'] = vl_col.apply(lambda r: extract_utr(str(r.get('particulars', '')) + ' ' + str(r.get('doc_no', ''))), axis=1)
    cl_col['utr'] = cl_col.apply(lambda r: extract_utr(str(r.get('doc_no', '')) + ' ' + str(r.get('doc_type', ''))), axis=1)

    for idx, vrow in vl_col.iterrows():
        matched = False
        basis = ''
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
            vl.at[idx, '_remark'] = 'Matched'
            vl.at[idx, '_match_ref'] = str(crow.get('doc_no', ''))
            cl.at[crow['_idx'], '_matched'] = True
            cl.at[crow['_idx'], '_remark'] = 'Matched'
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
                'Remark': 'Matched',
            })

    # ── STEP 5: Mark remaining unmatched ──
    unmatched_vl = vl[~vl['_matched']]
    unmatched_cl = cl[~cl['_matched']]

    for idx, r in unmatched_vl.iterrows():
        vl.at[idx, '_remark'] = 'Unmatched'
        entry = {
            'Doc No': r.get('doc_no', ''),
            'Date': r.get('doc_date', ''),
            'Type': r.get('doc_type', ''),
            'Particulars': r.get('particulars', ''),
            'Debit': r.get('debit', 0),
            'Credit': r.get('credit', 0),
            'Source': 'Vendor Ledger',
            'Remark': 'Unmatched',
        }
        if is_debit_note(r.get('doc_type', '')):
            results['dn_unmatched_vl'].append(entry)
        elif is_collection(r.get('doc_type', '')):
            results['collection_unmatched_vl'].append(entry)
        else:
            results['invoice_unmatched_vl'].append(entry)

    for idx, r in unmatched_cl.iterrows():
        cl.at[idx, '_remark'] = 'Unmatched'
        entry = {
            'Doc No': r.get('doc_no', ''),
            'Date': r.get('doc_date', ''),
            'Type': r.get('doc_type', ''),
            'Debit': r.get('debit', 0),
            'Credit': r.get('credit', 0),
            'Source': 'Customer Ledger',
            'Remark': 'Unmatched',
        }
        if is_debit_note(r.get('doc_type', '')):
            results['dn_unmatched_cl'].append(entry)
        elif is_collection(r.get('doc_type', '')):
            results['collection_unmatched_cl'].append(entry)
        else:
            results['invoice_unmatched_cl'].append(entry)

    # Store annotated ledgers for the "Ledger with Remarks" sheets
    results['vl_annotated'] = vl
    results['cl_annotated'] = cl

    return results

# ─────────────────────────────────────────────
# EXCEL EXPORT — with Vendor & Customer Ledger tabs + Remarks
# ─────────────────────────────────────────────

def build_excel(results, vl_orig, cl_orig):
    output = BytesIO()
    wb = openpyxl.Workbook()

    COLORS = {
        'matched_fill': 'C6EFCE',       # green
        'unmatched_fill': 'FFC7CE',     # red
        'reversal_fill': 'FFEB9C',      # orange/yellow
        'header_dark': '1C2130',
        'header_green': '1A6B45',
        'header_red': 'A32035',
        'header_orange': 'B85C00',
        'header_blue': '1A3A6B',
        'alt_row': 'F8F9FB',
        'border': 'D0D5E0',
    }

    thin = Side(style='thin', color=COLORS['border'])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, headers, row=1, color='1C2130'):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
            cell.fill = PatternFill(fill_type='solid', fgColor=color)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        ws.row_dimensions[row].height = 30

    def style_row(ws, row_num, ncols, fill_color):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.fill = PatternFill(fill_type='solid', fgColor=fill_color)
            cell.border = border
            cell.alignment = Alignment(vertical='center')
            cell.font = Font(name='Calibri', size=9)

    def auto_width(ws, min_w=10, max_w=45):
        for col in ws.columns:
            max_len = 0
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value or '')))
                except:
                    pass
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w, max(min_w, max_len + 2))

    def fmt_date(val):
        try:
            if pd.isna(val):
                return ''
            return pd.to_datetime(val).strftime('%d-%b-%Y')
        except:
            return str(val)

    def write_cell(ws, row, col, val):
        cell = ws.cell(row=row, column=col)
        if isinstance(val, (pd.Timestamp, datetime)):
            cell.value = fmt_date(val)
        elif isinstance(val, float) and not pd.isna(val):
            cell.value = round(val, 2)
            cell.number_format = '#,##0.00'
        elif pd.isna(val) if not isinstance(val, str) else False:
            cell.value = ''
        else:
            cell.value = val
        return cell

    # ══════════════════════════════════════════
    # SHEET 1: SUMMARY
    # ══════════════════════════════════════════
    ws_sum = wb.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    # Title
    ws_sum.merge_cells('A1:H1')
    title_cell = ws_sum['A1']
    title_cell.value = '⚖️ VENDOR RECONCILIATION SUMMARY'
    title_cell.font = Font(bold=True, size=14, color='FFFFFF', name='Calibri')
    title_cell.fill = PatternFill(fill_type='solid', fgColor='1C2130')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[1].height = 35

    ws_sum['A2'].value = f'Generated on: {datetime.now().strftime("%d-%b-%Y %H:%M")}'
    ws_sum['A2'].font = Font(italic=True, size=9, color='888888')
    ws_sum.row_dimensions[2].height = 18

    # Section: Reconciliation Overview
    headers = ['Category', 'Total (VL)', 'Matched', 'Unmatched (VL)', 'Total (CL)', 'Matched', 'Unmatched (CL)', 'Remarks']
    style_header(ws_sum, headers, row=4, color='1A3A6B')

    inv_vl_total = len(results['invoice_matched']) + len(results['invoice_unmatched_vl'])
    inv_cl_total = len(results['invoice_matched']) + len(results['invoice_unmatched_cl'])
    dn_vl_total = len(results['dn_matched']) + len(results['dn_unmatched_vl'])
    dn_cl_total = len(results['dn_matched']) + len(results['dn_unmatched_cl'])
    col_vl_total = len(results['collection_matched']) + len(results['collection_unmatched_vl'])
    col_cl_total = len(results['collection_matched']) + len(results['collection_unmatched_cl'])

    summary_rows = [
        ['Invoices', inv_vl_total, len(results['invoice_matched']), len(results['invoice_unmatched_vl']),
         inv_cl_total, len(results['invoice_matched']), len(results['invoice_unmatched_cl']), 'Matched by Document Number'],
        ['Debit Notes', dn_vl_total, len(results['dn_matched']), len(results['dn_unmatched_vl']),
         dn_cl_total, len(results['dn_matched']), len(results['dn_unmatched_cl']), 'Matched by Doc No / Period+Amount'],
        ['Collections', col_vl_total, len(results['collection_matched']), len(results['collection_unmatched_vl']),
         col_cl_total, len(results['collection_matched']), len(results['collection_unmatched_cl']), 'Matched by UTR / Period+Amount'],
        ['Reversals (VL)', len(results['reversal_matched']) * 2 + len(results['reversal_unmatched']),
         len(results['reversal_matched']), len(results['reversal_unmatched']),
         '-', '-', '-', 'Complete Reversal / Saleable Return'],
    ]

    row_colors = ['FFFFFF', 'F8F9FB', 'FFFFFF', 'FFF9E6']
    for r_idx, row_data in enumerate(summary_rows, 5):
        fill = row_colors[r_idx - 5]
        for c_idx, val in enumerate(row_data, 1):
            cell = ws_sum.cell(row=r_idx, column=c_idx, value=val)
            cell.fill = PatternFill(fill_type='solid', fgColor=fill)
            cell.border = border
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(vertical='center')
            if c_idx in [3, 6]:  # matched cols - green
                cell.font = Font(name='Calibri', size=10, color='1A6B45', bold=True)
            if c_idx in [4, 7]:  # unmatched cols - red
                cell.font = Font(name='Calibri', size=10, color='A32035', bold=True)
        ws_sum.row_dimensions[r_idx].height = 22

    # Total row
    total_row = r_idx + 1
    totals = [
        'TOTAL',
        inv_vl_total + dn_vl_total + col_vl_total,
        len(results['invoice_matched']) + len(results['dn_matched']) + len(results['collection_matched']),
        len(results['invoice_unmatched_vl']) + len(results['dn_unmatched_vl']) + len(results['collection_unmatched_vl']),
        inv_cl_total + dn_cl_total + col_cl_total,
        len(results['invoice_matched']) + len(results['dn_matched']) + len(results['collection_matched']),
        len(results['invoice_unmatched_cl']) + len(results['dn_unmatched_cl']) + len(results['collection_unmatched_cl']),
        '',
    ]
    for c_idx, val in enumerate(totals, 1):
        cell = ws_sum.cell(row=total_row, column=c_idx, value=val)
        cell.fill = PatternFill(fill_type='solid', fgColor='1C2130')
        cell.font = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[total_row].height = 25

    # Match % row
    pct_row = total_row + 1
    vl_match_pct = round(totals[2] / totals[1] * 100, 1) if totals[1] else 0
    cl_match_pct = round(totals[5] / totals[4] * 100, 1) if totals[4] else 0
    ws_sum.cell(row=pct_row, column=1, value='Match Rate')
    ws_sum.cell(row=pct_row, column=3, value=f'{vl_match_pct}%')
    ws_sum.cell(row=pct_row, column=6, value=f'{cl_match_pct}%')
    for c in [1, 3, 6]:
        cell = ws_sum.cell(row=pct_row, column=c)
        cell.font = Font(bold=True, color='1A6B45', name='Calibri', size=11)
        cell.fill = PatternFill(fill_type='solid', fgColor='E8F5EE')
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[pct_row].height = 22

    auto_width(ws_sum)

    # ══════════════════════════════════════════
    # SHEET 2: VENDOR LEDGER WITH REMARKS
    # ══════════════════════════════════════════
    vl_ann = results['vl_annotated']
    ws_vl = wb.create_sheet('Vendor Ledger')
    ws_vl.sheet_view.showGridLines = False
    ws_vl.freeze_panes = 'A2'

    # Choose display columns
    vl_display_cols = ['doc_date', 'doc_no', 'doc_type', 'particulars', 'debit', 'credit']
    vl_display_cols = [c for c in vl_display_cols if c in vl_ann.columns]
    vl_display_cols += ['_remark', '_match_ref']

    vl_headers = {
        'doc_date': 'Doc Date', 'doc_no': 'Doc No', 'doc_type': 'Doc Type',
        'particulars': 'Particulars', 'debit': 'Debit', 'credit': 'Credit',
        '_remark': 'Remark', '_match_ref': 'Matched With'
    }
    headers = [vl_headers.get(c, c) for c in vl_display_cols]
    style_header(ws_vl, headers, row=1, color='1C2130')

    for r_idx, (_, row) in enumerate(vl_ann[vl_display_cols].iterrows(), 2):
        remark = str(row.get('_remark', ''))
        if 'Unmatched' in remark:
            fill = COLORS['unmatched_fill']
        elif 'Reversal' in remark or 'Reversed' in remark:
            fill = COLORS['reversal_fill']
        else:
            fill = COLORS['matched_fill']

        for c_idx, col in enumerate(vl_display_cols, 1):
            val = row[col]
            cell = write_cell(ws_vl, r_idx, c_idx, val)
            cell.fill = PatternFill(fill_type='solid', fgColor=fill)
            cell.border = border
            cell.alignment = Alignment(vertical='center')
            cell.font = Font(name='Calibri', size=9)
        ws_vl.row_dimensions[r_idx].height = 18

    auto_width(ws_vl)

    # ══════════════════════════════════════════
    # SHEET 3: CUSTOMER LEDGER WITH REMARKS
    # ══════════════════════════════════════════
    cl_ann = results['cl_annotated']
    ws_cl = wb.create_sheet('Customer Ledger')
    ws_cl.sheet_view.showGridLines = False
    ws_cl.freeze_panes = 'A2'

    cl_display_cols = ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit']
    cl_display_cols = [c for c in cl_display_cols if c in cl_ann.columns]
    cl_display_cols += ['_remark', '_match_ref']

    cl_headers = {
        'doc_date': 'Doc Date', 'doc_no': 'Doc No', 'doc_type': 'Doc Type',
        'debit': 'Debit (LC)', 'credit': 'Credit (LC)',
        '_remark': 'Remark', '_match_ref': 'Matched With'
    }
    headers = [cl_headers.get(c, c) for c in cl_display_cols]
    style_header(ws_cl, headers, row=1, color='1C2130')

    for r_idx, (_, row) in enumerate(cl_ann[cl_display_cols].iterrows(), 2):
        remark = str(row.get('_remark', ''))
        fill = COLORS['unmatched_fill'] if 'Unmatched' in remark else COLORS['matched_fill']

        for c_idx, col in enumerate(cl_display_cols, 1):
            val = row[col]
            cell = write_cell(ws_cl, r_idx, c_idx, val)
            cell.fill = PatternFill(fill_type='solid', fgColor=fill)
            cell.border = border
            cell.alignment = Alignment(vertical='center')
            cell.font = Font(name='Calibri', size=9)
        ws_cl.row_dimensions[r_idx].height = 18

    auto_width(ws_cl)

    # ══════════════════════════════════════════
    # REMAINING DETAIL SHEETS
    # ══════════════════════════════════════════
    def write_sheet(title, data, color):
        if not data:
            return
        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'A2'
        df = pd.DataFrame(data)
        headers = list(df.columns)
        style_header(ws, headers, row=1, color=color)
        for r, (_, row) in enumerate(df.iterrows(), 2):
            for c, val in enumerate(row.values, 1):
                cell = write_cell(ws, r, c, val)
                fill = 'F8F9FB' if r % 2 == 0 else 'FFFFFF'
                cell.fill = PatternFill(fill_type='solid', fgColor=fill)
                cell.border = border
                cell.alignment = Alignment(vertical='center')
                cell.font = Font(name='Calibri', size=9)
        auto_width(ws)

    write_sheet('Inv - Matched', results['invoice_matched'], '1A6B45')
    write_sheet('Inv - Unmatched VL', results['invoice_unmatched_vl'], 'A32035')
    write_sheet('Inv - Unmatched CL', results['invoice_unmatched_cl'], 'A32035')
    write_sheet('DN - Matched', results['dn_matched'], '1A6B45')
    write_sheet('DN - Unmatched VL', results['dn_unmatched_vl'], 'A32035')
    write_sheet('DN - Unmatched CL', results['dn_unmatched_cl'], 'A32035')
    write_sheet('Collections - Matched', results['collection_matched'], '1A6B45')
    write_sheet('Collections - Unmatch VL', results['collection_unmatched_vl'], 'A32035')
    write_sheet('Collections - Unmatch CL', results['collection_unmatched_cl'], 'A32035')
    write_sheet('Reversals - Matched', results['reversal_matched'], 'B85C00')
    write_sheet('Reversals - Unmatched', results['reversal_unmatched'], 'A32035')

    wb.save(output)
    return output.getvalue()

# ─────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────

def display_df(df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty) or (isinstance(df, list) and len(df) == 0):
        st.info("No records in this category.")
        return
    if isinstance(df, list):
        df = pd.DataFrame(df)
    df = df.copy()
    for col in df.columns:
        if 'date' in col.lower() or 'Date' in col:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%b-%Y').fillna('')
    st.dataframe(df, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────

def main():
    st.markdown("""
    <div class="recon-header">
        <div>
            <div class="recon-logo">⚖️ VendorSync</div>
            <div class="recon-subtitle">Vendor · Customer Ledger Reconciliation · For Indian CAs &amp; CFOs</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

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
        st.markdown("""
        <div class="info-box">
        1. Invoices → Doc Number<br>
        2. Reversals → Complete Reversal / Saleable Return (via Particulars)<br>
        3. Debit Notes → Doc No → Period+Amount<br>
        4. Collections → UTR → Period+Amount<br>
        5. All items get ✅ Matched / ❌ Unmatched remark
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("""
        <div class="info-box" style="font-size:0.72rem;">
        <b>Excel Output includes:</b><br>
        • Summary sheet with match rates<br>
        • Vendor Ledger tab with remarks<br>
        • Customer Ledger tab with remarks<br>
        • Detail sheets per category
        </div>
        """, unsafe_allow_html=True)

    # File upload
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<span class="section-tag tag-blue">VENDOR LEDGER</span>', unsafe_allow_html=True)
        vl_file = st.file_uploader("Upload Vendor Ledger (.xlsx / .xls)", type=['xlsx', 'xls'], key='vl')
    with col2:
        st.markdown('<span class="section-tag tag-blue">CUSTOMER LEDGER</span>', unsafe_allow_html=True)
        cl_file = st.file_uploader("Upload Customer Ledger (.xlsx / .xls)", type=['xlsx', 'xls'], key='cl')

    if not vl_file or not cl_file:
        st.markdown("""
        <div class="info-box" style="margin-top:2rem; border-left-color: #4f8eff;">
        📂 Upload both ledger files above to begin reconciliation. The engine will automatically detect columns,
        handle reversals (Complete Reversal / Saleable Return), match invoices, debit notes, and collections,
        and add Matched/Unmatched remarks to each row in the output Excel.
        </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Parsing ledgers..."):
        try:
            vl = load_vendor_ledger(vl_file)
            cl = load_customer_ledger(cl_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            return

    st.success(f"✅ Vendor Ledger: **{len(vl)}** rows  ·  Customer Ledger: **{len(cl)}** rows")

    with st.expander("👁 Preview Parsed Data"):
        pc1, pc2 = st.columns(2)
        with pc1:
            st.markdown("**Vendor Ledger (first 10 rows)**")
            display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'particulars', 'debit', 'credit'] if c in vl.columns]
            st.dataframe(vl[display_cols].head(10), use_container_width=True, hide_index=True)
        with pc2:
            st.markdown("**Customer Ledger (first 10 rows)**")
            display_cols = [c for c in ['doc_date', 'doc_no', 'doc_type', 'debit', 'credit'] if c in cl.columns]
            st.dataframe(cl[display_cols].head(10), use_container_width=True, hide_index=True)

    if st.button("▶ Run Reconciliation", use_container_width=False):
        with st.spinner("Running reconciliation engine..."):
            results = run_reconciliation(vl, cl, tolerance=tolerance)
        st.session_state['results'] = results
        st.session_state['vl'] = vl
        st.session_state['cl'] = cl

    if 'results' not in st.session_state:
        return

    results = st.session_state['results']
    vl = st.session_state['vl']
    cl = st.session_state['cl']

    # ── SUMMARY STATS ──
    total_matched = len(results['invoice_matched']) + len(results['dn_matched']) + len(results['collection_matched'])
    total_unmatched_vl = len(results['invoice_unmatched_vl']) + len(results['dn_unmatched_vl']) + len(results['collection_unmatched_vl'])
    total_unmatched_cl = len(results['invoice_unmatched_cl']) + len(results['dn_unmatched_cl']) + len(results['collection_unmatched_cl'])
    total_reversals = len(results['reversal_matched'])

    st.markdown(f"""
    <div class="stat-grid">
        <div class="stat-card matched">
            <div class="stat-label">Total Matched</div>
            <div class="stat-value">{total_matched}</div>
            <div class="stat-sub">Invoices + DN + Collections</div>
        </div>
        <div class="stat-card unmatched">
            <div class="stat-label">Unmatched (VL)</div>
            <div class="stat-value">{total_unmatched_vl}</div>
            <div class="stat-sub">CL Unmatched: {total_unmatched_cl}</div>
        </div>
        <div class="stat-card partial">
            <div class="stat-label">Reversals Reconciled</div>
            <div class="stat-value">{total_reversals}</div>
            <div class="stat-sub">Complete Reversal / Saleable Return</div>
        </div>
        <div class="stat-card total">
            <div class="stat-label">Invoice Matches</div>
            <div class="stat-value">{len(results['invoice_matched'])}</div>
            <div class="stat-sub">DN: {len(results['dn_matched'])} · Coll: {len(results['collection_matched'])}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── DETAILED SUMMARY TABLE ──
    st.markdown("### 📊 Reconciliation Summary")

    inv_vl = len(results['invoice_matched']) + len(results['invoice_unmatched_vl'])
    inv_cl = len(results['invoice_matched']) + len(results['invoice_unmatched_cl'])
    dn_vl = len(results['dn_matched']) + len(results['dn_unmatched_vl'])
    dn_cl = len(results['dn_matched']) + len(results['dn_unmatched_cl'])
    col_vl = len(results['collection_matched']) + len(results['collection_unmatched_vl'])
    col_cl = len(results['collection_matched']) + len(results['collection_unmatched_cl'])

    inv_pct = round(len(results['invoice_matched']) / inv_vl * 100, 1) if inv_vl else 0
    dn_pct = round(len(results['dn_matched']) / dn_vl * 100, 1) if dn_vl else 0
    col_pct = round(len(results['collection_matched']) / col_vl * 100, 1) if col_vl else 0

    st.markdown(f"""
    <table class="summary-table">
        <thead>
            <tr>
                <th>Category</th>
                <th class="num">VL Total</th>
                <th class="num">Matched</th>
                <th class="num">VL Unmatched</th>
                <th class="num">CL Total</th>
                <th class="num">CL Unmatched</th>
                <th class="num">Match %</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>🧾 Invoices</td>
                <td class="num">{inv_vl}</td>
                <td class="num matched-val">{len(results['invoice_matched'])}</td>
                <td class="num unmatched-val">{len(results['invoice_unmatched_vl'])}</td>
                <td class="num">{inv_cl}</td>
                <td class="num unmatched-val">{len(results['invoice_unmatched_cl'])}</td>
                <td class="num matched-val">{inv_pct}%</td>
            </tr>
            <tr>
                <td>📝 Debit Notes</td>
                <td class="num">{dn_vl}</td>
                <td class="num matched-val">{len(results['dn_matched'])}</td>
                <td class="num unmatched-val">{len(results['dn_unmatched_vl'])}</td>
                <td class="num">{dn_cl}</td>
                <td class="num unmatched-val">{len(results['dn_unmatched_cl'])}</td>
                <td class="num matched-val">{dn_pct}%</td>
            </tr>
            <tr>
                <td>💰 Collections</td>
                <td class="num">{col_vl}</td>
                <td class="num matched-val">{len(results['collection_matched'])}</td>
                <td class="num unmatched-val">{len(results['collection_unmatched_vl'])}</td>
                <td class="num">{col_cl}</td>
                <td class="num unmatched-val">{len(results['collection_unmatched_cl'])}</td>
                <td class="num matched-val">{col_pct}%</td>
            </tr>
            <tr>
                <td>🔁 Reversals</td>
                <td class="num">{len(results['reversal_matched']) * 2 + len(results['reversal_unmatched'])}</td>
                <td class="num matched-val">{len(results['reversal_matched'])}</td>
                <td class="num unmatched-val">{len(results['reversal_unmatched'])}</td>
                <td class="num">—</td>
                <td class="num">—</td>
                <td class="num">—</td>
            </tr>
            <tr class="total-row">
                <td><b>TOTAL</b></td>
                <td class="num"><b>{inv_vl + dn_vl + col_vl}</b></td>
                <td class="num matched-val"><b>{total_matched}</b></td>
                <td class="num unmatched-val"><b>{total_unmatched_vl}</b></td>
                <td class="num"><b>{inv_cl + dn_cl + col_cl}</b></td>
                <td class="num unmatched-val"><b>{total_unmatched_cl}</b></td>
                <td class="num matched-val"><b>—</b></td>
            </tr>
        </tbody>
    </table>
    """, unsafe_allow_html=True)

    # ── DOWNLOAD ──
    excel_data = build_excel(results, vl, cl)
    st.download_button(
        label="⬇ Download Full Reconciliation Report (.xlsx)",
        data=excel_data,
        file_name=f"VendorSync_Reconciliation_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

    # ── DETAIL TABS ──
    tabs = st.tabs([
        "🧾 Invoices",
        "📝 Debit Notes",
        "💰 Collections",
        "🔁 Reversals",
        "⚠️ All Unmatched",
    ])

    with tabs[0]:
        st.markdown('<span class="section-tag tag-matched">MATCHED INVOICES</span>', unsafe_allow_html=True)
        display_df(results['invoice_matched'])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — VENDOR LEDGER</span>', unsafe_allow_html=True)
            display_df(results['invoice_unmatched_vl'])
        with c2:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — CUSTOMER LEDGER</span>', unsafe_allow_html=True)
            display_df(results['invoice_unmatched_cl'])

    with tabs[1]:
        st.markdown('<span class="section-tag tag-matched">MATCHED DEBIT NOTES</span>', unsafe_allow_html=True)
        display_df(results['dn_matched'])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — VENDOR LEDGER</span>', unsafe_allow_html=True)
            display_df(results['dn_unmatched_vl'])
        with c2:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — CUSTOMER LEDGER</span>', unsafe_allow_html=True)
            display_df(results['dn_unmatched_cl'])

    with tabs[2]:
        st.markdown('<span class="section-tag tag-matched">MATCHED COLLECTIONS</span>', unsafe_allow_html=True)
        display_df(results['collection_matched'])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — VENDOR LEDGER</span>', unsafe_allow_html=True)
            display_df(results['collection_unmatched_vl'])
        with c2:
            st.markdown('<span class="section-tag tag-unmatched">UNMATCHED — CUSTOMER LEDGER</span>', unsafe_allow_html=True)
            display_df(results['collection_unmatched_cl'])

    with tabs[3]:
        st.markdown('<span class="section-tag tag-partial">REVERSALS RECONCILED (Complete Reversal / Saleable Return)</span>', unsafe_allow_html=True)
        st.caption("Reversal invoices matched to their original invoice via the Particulars column.")
        display_df(results['reversal_matched'])
        if results['reversal_unmatched']:
            st.markdown('<span class="section-tag tag-unmatched">REVERSALS — ORIGINAL NOT FOUND</span>', unsafe_allow_html=True)
            display_df(results['reversal_unmatched'])

    with tabs[4]:
        all_unmatched = []
        for item in results['invoice_unmatched_vl'] + results['dn_unmatched_vl'] + results['collection_unmatched_vl']:
            item = dict(item); item['Ledger'] = 'Vendor'; all_unmatched.append(item)
        for item in results['invoice_unmatched_cl'] + results['dn_unmatched_cl'] + results['collection_unmatched_cl']:
            item = dict(item); item['Ledger'] = 'Customer'; all_unmatched.append(item)
        st.markdown('<span class="section-tag tag-unmatched">ALL UNMATCHED ITEMS</span>', unsafe_allow_html=True)
        display_df(all_unmatched)


if __name__ == "__main__":
    main()
