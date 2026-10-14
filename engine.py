"""
engine.py — MMU Hub Report computation + validation engine.
Implements Build Spec v10, generalized to accept uploaded files (any subset
may be missing) instead of fixed local paths, with graceful per-parameter
degradation instead of an all-or-nothing failure.

Division-agnostic: sub-divisions, the division name and the office counts are
all derived from the uploaded Master file, so any Division that supplies a
Master in the standard format gets the same report layout.

Only Raw Data carries live formulas (with cached values written alongside,
via xlsxwriter's write_formula(..., value=...), so the numbers show up
immediately without needing a LibreOffice/Excel recalculation pass — this
keeps the app self-contained for Streamlit Cloud deployment). Every other
visible sheet is hardcoded pre-computed values, exactly as in the desktop
build.
"""
import io
import os
import re
import numpy as np
import pandas as pd
import xlsxwriter  # noqa: F401 (re-exported for convenience elsewhere)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
MASTER_PATH = os.path.join(DATA_DIR, 'master_file.xlsx')
LOGO_PATH = os.path.join(DATA_DIR, 'logo.png')

KNOWN_OFFICE_TYPES = {'BPO', 'SPO', 'HPO'}


def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def save_master_file(uploaded_file):
    """Persist an uploaded Master file to local disk, overwriting any previous one."""
    ensure_data_dir()
    with open(MASTER_PATH, 'wb') as fh:
        fh.write(uploaded_file.getbuffer())


def load_persisted_master():
    """Return the persisted Master file path if one exists, else None."""
    return MASTER_PATH if os.path.exists(MASTER_PATH) else None


def save_logo_file(uploaded_file):
    ensure_data_dir()
    ext = os.path.splitext(uploaded_file.name)[1].lower() or '.png'
    path = os.path.join(DATA_DIR, f'logo{ext}')
    with open(path, 'wb') as fh:
        fh.write(uploaded_file.getbuffer())
    return path


def load_persisted_logo():
    if not os.path.isdir(DATA_DIR):
        return None
    for ext in ('.png', '.jpg', '.jpeg', '.svg'):
        p = os.path.join(DATA_DIR, f'logo{ext}')
        if os.path.exists(p):
            return p
    return None


# ----------------------------------------------------------------- I/O ---

def read_csv_any(fileobj):
    """Read a Streamlit UploadedFile (or path) as a DataFrame. Tolerates the
    common non-UTF-8 encodings exports sometimes use, and rewinds the buffer so
    the same upload can be re-read on a Streamlit rerun."""
    last = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            if hasattr(fileobj, 'seek'):
                fileobj.seek(0)
            return pd.read_csv(fileobj, encoding=enc)
        except UnicodeDecodeError as e:
            last = e
    raise last


def read_many_csv(fileobjs):
    """Concatenate zero or more uploaded CSV parts (e.g. DP Range 1 + Range 2)
    into a single DataFrame. Returns None if the list is empty."""
    if not fileobjs:
        return None
    dfs = [read_csv_any(f) for f in fileobjs]
    return pd.concat(dfs, ignore_index=True)


# ------------------------------------------------------------- findings ---
# A "finding" is {'level': 'pass'|'info'|'warning'|'error', 'check': str, 'detail': str}.
# Errors block report generation; warnings are shown and logged but do not.

def _f(level, check, detail=''):
    return {'level': level, 'check': check, 'detail': detail}


def has_errors(findings):
    return any(x['level'] == 'error' for x in findings)


def _preview(values, n=8):
    values = list(values)
    s = ', '.join(str(v) for v in values[:n])
    return s + (f' … (+{len(values) - n} more)' if len(values) > n else '')


# -------------------------------------------------------------- Master ---

MASTER_COLS = {  # Master header -> internal name   (True = required)
    'Office ID': ('office_id', True),
    'Office Name': ('office_name', True),
    'Sub Division Name': ('sub_division', True),
    'Office Type Code': ('office_type', True),
    'Sub Office Name': ('sub_office', False),
}
MASTER_OPTIONAL_INFO = {'Division Office Name': 'division_name',
                        'Region Office Name': 'region_name', 'Circle Name': 'circle_name'}


def read_master_full(fileobj):
    """Read + validate a Master file. Returns (master_df, info, findings).
    master_df is None when there are blocking errors. info carries the division /
    region / circle names and the office counts used for headers and labels."""
    findings = []
    try:
        if hasattr(fileobj, 'seek'):
            fileobj.seek(0)
        raw = pd.read_excel(fileobj)
    except Exception as e:  # unreadable / not an xlsx
        return None, {}, [_f('error', 'Master: file readable',
                             f'Could not open the Master as an Excel (.xlsx) file: {e}')]

    raw.columns = [str(c).strip() for c in raw.columns]
    lower = {c.lower(): c for c in raw.columns}
    rename, missing = {}, []
    for head, (internal, required) in MASTER_COLS.items():
        src = lower.get(head.lower())
        if src is None:
            if required:
                missing.append(head)
        else:
            rename[src] = internal
    if missing:
        return None, {}, [_f('error', 'Master: required columns',
                             f"Missing column(s): {', '.join(missing)}. The Master must contain the standard "
                             f"headers: Office ID, Office Name, Sub Division Name, Office Type Code "
                             f"(and normally Sub Office Name). Columns found: {_preview(raw.columns, 20)}")]

    m = raw.rename(columns=rename)
    if 'sub_office' not in m.columns:
        m['sub_office'] = ''
        findings.append(_f('warning', 'Master: Sub Office Name column', 'Column not found — left blank.'))
    info = {}
    for head, internal in MASTER_OPTIONAL_INFO.items():
        src = lower.get(head.lower())
        vals = raw[src].dropna().astype(str).str.strip() if src else pd.Series(dtype=str)
        vals = vals[vals != '']
        info[internal] = vals.mode().iloc[0] if len(vals) else ''
        if internal == 'division_name' and vals.nunique() > 1:
            findings.append(_f('warning', 'Master: single division',
                               f'Master lists more than one Division ({_preview(vals.unique())}). '
                               f'A report is built for one Division — check you uploaded the right Master.'))
    m = m[['office_id', 'office_name', 'sub_office', 'sub_division', 'office_type']].copy()
    m = m.dropna(how='all')

    if len(m) == 0:
        return None, info, [_f('error', 'Master: not empty', 'The Master has no rows.')]

    ids = pd.to_numeric(m['office_id'], errors='coerce')
    bad = ids.isna()
    if bad.any():
        return None, info, [_f('error', 'Master: Office ID values',
                               f'{int(bad.sum())} row(s) have a blank or non-numeric Office ID '
                               f'(e.g. office: {_preview(m.loc[bad, "office_name"].tolist(), 5)}).')]
    m['office_id'] = ids.astype('int64')

    for c in ('office_name', 'sub_office', 'sub_division', 'office_type'):
        m[c] = m[c].fillna('').astype(str).str.strip()
    m['office_type'] = m['office_type'].str.upper()

    dup = m[m['office_id'].duplicated(keep=False)]
    if len(dup):
        ids_dup = sorted(dup['office_id'].unique())
        return None, info, [_f('error', 'Master: unique Office IDs',
                               f'{len(ids_dup)} Office ID(s) appear more than once: {_preview(ids_dup)}. '
                               f'Every office must have exactly one row — duplicates would double-count figures.')]
    findings.append(_f('pass', 'Master: unique Office IDs', f'{len(m)} offices, no duplicate Office IDs'))

    blank_sd = (m['sub_division'] == '')
    if blank_sd.any():
        return None, info, [_f('error', 'Master: Sub Division Name',
                               f'{int(blank_sd.sum())} office(s) have no Sub Division Name: '
                               f'{_preview(m.loc[blank_sd, "office_name"].tolist(), 6)}.')]

    unknown = sorted(set(m['office_type']) - KNOWN_OFFICE_TYPES)
    if unknown:
        findings.append(_f('warning', 'Master: Office Type Code',
                           f'Unrecognised type code(s): {_preview(unknown)}. Expected BPO / SPO / HPO. '
                           f'Anything other than BPO is treated as having letter boxes.'))

    counts = m['office_type'].value_counts().to_dict()
    info['n_offices'] = len(m)
    info['type_counts'] = counts
    info['sub_divisions'] = sorted(m['sub_division'].unique())
    info['sub_division_counts'] = m['sub_division'].value_counts().to_dict()
    findings.append(_f('info', 'Master: composition',
                       f"{len(m)} offices; " + ', '.join(f'{k}={v}' for k, v in sorted(counts.items())) +
                       '; sub-divisions: ' + ', '.join(f'{k}={v}' for k, v in
                                                       sorted(info['sub_division_counts'].items()))))
    return m.reset_index(drop=True), info, findings


def read_master(fileobj):
    """Back-compat wrapper: returns the cleaned Master or raises ValueError."""
    m, _info, findings = read_master_full(fileobj)
    if m is None:
        raise ValueError('; '.join(x['detail'] for x in findings if x['level'] == 'error'))
    return m


def get_sub_divisions(df):
    """Sub-divisions of the Division, A→Z (derived from the Master — never hardcoded)."""
    return sorted(df['sub_division'].dropna().unique())


def default_record_prefix(division_name):
    """Office-record-number prefix, e.g. 'Karad Division' -> 'KRDDN'
    (consonants of the first word, up to 3, + 'DN'). Editable in the UI."""
    word = re.sub(r'[^A-Za-z]', '', (division_name or '').split(' ')[0]).upper()
    cons = ''.join(ch for ch in word if ch not in 'AEIOU')[:3]
    return (cons or word[:3] or 'DIV') + 'DN'


# --------------------------------------------------- input-file schemas ---

SCHEMAS = {
    'dp':  {'label': 'Delivery Productivity', 'id': 'office-id',
            'required': ['office-id', 'invoice-count', 'deposit-count'],
            'numeric': ['invoice-count', 'deposit-count'], 'signature': 'invoice-count'},
    'dss': {'label': 'DSS Usage', 'id': 'office_id',
            'required': ['office_id', 'total_pdm_art_count', 'total_dss_art_count'],
            'numeric': ['total_pdm_art_count', 'total_dss_art_count'], 'signature': 'total_pdm_art_count'},
    'lb':  {'label': 'Letter Box Clearance', 'id': 'office-id',
            'required': ['office-id', 'total-letterboxes', 'total-cleared'],
            'numeric': ['total-letterboxes', 'total-cleared'], 'signature': 'total-letterboxes'},
    'cod': {'label': 'COD Digital Transaction', 'id': 'Office ID',
            'required': ['Office ID', 'COD Digital Count', 'Total COD Count'],
            'numeric': ['COD Digital Count', 'Total COD Count'], 'signature': 'COD Digital Count'},
}


def detect_category(df):
    for key, s in SCHEMAS.items():
        if s['signature'] in df.columns:
            return key
    return None


def _validate_file(slot_label, cat, df, master_ids):
    """Schema + sanity checks for one uploaded file. Returns findings."""
    out = []
    if df is None:
        return out
    s = SCHEMAS[cat]
    if len(df) == 0:
        return [_f('error', f'{slot_label}: not empty', 'The file has no data rows.')]
    missing = [c for c in s['required'] if c not in df.columns]
    if missing:
        guess = detect_category(df)
        hint = (f" This looks like a {SCHEMAS[guess]['label']} file — it may be in the wrong slot."
                if guess and guess != cat else '')
        return [_f('error', f'{slot_label}: columns',
                   f"Expected a {s['label']} export but column(s) {', '.join(missing)} are missing.{hint}")]
    for c in s['numeric'] + [s['id']]:
        coerced = pd.to_numeric(df[c], errors='coerce')
        n_bad = int((coerced.isna() & df[c].notna()).sum())
        if n_bad:
            return [_f('error', f'{slot_label}: numeric columns',
                       f'Column "{c}" has {n_bad} non-numeric value(s) (e.g. {_preview(df.loc[coerced.isna() & df[c].notna(), c].tolist(), 3)}).')]
    out.append(_f('pass', f'{slot_label}: columns', f"{len(df)} rows, expected {s['label']} columns present"))

    ids = pd.to_numeric(df[s['id']], errors='coerce')
    for c in s['numeric']:
        n_neg = int((pd.to_numeric(df[c], errors='coerce') < 0).sum())
        if n_neg:
            out.append(_f('warning', f'{slot_label}: negative values', f'Column "{c}" has {n_neg} negative value(s).'))

    unmatched = sorted(set(ids.dropna().astype('int64')) - set(master_ids) - {0})
    if unmatched:
        namecol = next((c for c in ('office-name', 'office_name', 'Office Name') if c in df.columns), None)
        names = []
        for u in unmatched[:8]:
            nm = df.loc[ids == u, namecol].iloc[0] if namecol else ''
            names.append(f'{nm} ({u})' if nm else str(u))
        extra = f' … (+{len(unmatched) - 8} more)' if len(unmatched) > 8 else ''
        out.append(_f('warning', f'{slot_label}: offices not in Master',
                      f'{len(unmatched)} office ID(s) in this file are not in the Master and are EXCLUDED from the '
                      f'report: {", ".join(names)}{extra}. If these are new offices, update the Master.'))
    return out


def _totals_by_office(df, cat):
    s = SCHEMAS[cat]
    col = {'dp': 'invoice-count', 'dss': 'total_pdm_art_count', 'cod': 'Total COD Count',
           'lb': 'total-cleared'}[cat]
    d = _drop_exact_dupes(df) if cat in ('dp', 'dss') else df.copy()
    d[s['id']] = pd.to_numeric(d[s['id']], errors='coerce')
    d = d[d[s['id']].notna() & (d[s['id']] != 0)]
    return d.groupby(s['id'])[col].sum()


def run_validations(master, info, frames, single_date=None, range_start=None, range_end=None):
    """Full pre-generation validation. `frames` maps slot key -> DataFrame or None, with keys
    dp_r1 dp_r2 dp_sd dss_r dss_sd lb_r lb_sd cod_r cod_sd. Returns findings."""
    out = []
    mids = set(master['office_id'])
    slots = [('dp_r1', 'dp', 'Delivery Productivity — Range 1'), ('dp_r2', 'dp', 'Delivery Productivity — Range 2'),
             ('dp_sd', 'dp', 'Delivery Productivity — Single Date'), ('dss_r', 'dss', 'DSS Usage — Range'),
             ('dss_sd', 'dss', 'DSS Usage — Single Date'), ('lb_r', 'lb', 'Letter Box Clearance — Range'),
             ('lb_sd', 'lb', 'Letter Box Clearance — Single Date'), ('cod_r', 'cod', 'COD Digital Transaction — Range'),
             ('cod_sd', 'cod', 'COD Digital Transaction — Single Date')]
    for key, cat, label in slots:
        out += _validate_file(label, cat, frames.get(key), mids)
    if has_errors(out):
        return out  # no point cross-checking files whose structure is wrong

    # one-row-per-office expectations
    for key, cat, label in slots:
        d = frames.get(key)
        if d is None or key in ('lb_r', 'cod_r', 'cod_sd'):
            continue
        idc = SCHEMAS[cat]['id']
        n_exact = int(d.duplicated().sum())
        if n_exact:
            out.append(_f('warning', f'{label}: repeated rows',
                          f'{n_exact} row(s) are exact repeats of another row and are IGNORED (counted once). '
                          f'Check the export was not run or pasted twice.'))
        d = d.drop_duplicates()
        ids = pd.to_numeric(d[idc], errors='coerce')
        ids = ids[ids.notna() & (ids != 0)]
        n_dup = int(ids.duplicated().sum())
        if n_dup:
            out.append(_f('warning', f'{label}: office on several rows',
                          f'{n_dup} office(s) appear on more than one row with different figures '
                          f'({_preview(sorted(set(ids[ids.duplicated()].astype("int64"))), 6)}); their figures are '
                          f'added together. This file is expected to have one row per office.'))

    # percentage-bound sanity on raw counts
    for key in ('dp_r1', 'dp_r2', 'dp_sd'):
        d = frames.get(key)
        if d is not None:
            n = int((pd.to_numeric(d['deposit-count'], errors='coerce') > pd.to_numeric(d['invoice-count'], errors='coerce')).sum())
            if n:
                out.append(_f('warning', f'{dict((k, l) for k, _, l in slots)[key]}: deposit > invoice',
                              f'{n} office(s) have more articles in deposit than invoiced — Delivery % would be negative.'))
    for key in ('dss_r', 'dss_sd'):
        d = frames.get(key)
        if d is not None:
            n = int((pd.to_numeric(d['total_dss_art_count'], errors='coerce') > pd.to_numeric(d['total_pdm_art_count'], errors='coerce')).sum())
            if n:
                out.append(_f('warning', f'{dict((k, l) for k, _, l in slots)[key]}: DSS > PDM',
                              f'{n} office(s) have more DSS articles than PDM articles — DSS % would exceed 100%.'))
    for key in ('cod_r', 'cod_sd'):
        d = frames.get(key)
        if d is not None:
            n = int((pd.to_numeric(d['COD Digital Count'], errors='coerce') > pd.to_numeric(d['Total COD Count'], errors='coerce')).sum())
            if n:
                out.append(_f('warning', f'{dict((k, l) for k, _, l in slots)[key]}: digital > total',
                              f'{n} row(s) have a COD Digital Count above Total COD Count.'))
    for key in ('lb_r', 'lb_sd'):
        d = frames.get(key)
        if d is not None:
            n = int((pd.to_numeric(d['total-cleared'], errors='coerce') > pd.to_numeric(d['total-letterboxes'], errors='coerce')).sum())
            if n:
                out.append(_f('warning', f'{dict((k, l) for k, _, l in slots)[key]}: cleared > boxes',
                              f'{n} row(s) have more boxes cleared than boxes — LB % would exceed 100%.'))

    # Range vs Single Date consistency (catches swapped / mis-slotted files)
    range_days = None
    if range_start and range_end:
        range_days = (range_end - range_start).days + 1
        if single_date and not (range_start <= single_date <= range_end):
            out.append(_f('warning', 'Dates: Single Date within Range',
                          f'Single Date {single_date:%d.%m.%Y} falls outside the Range {range_start:%d.%m.%Y} – '
                          f'{range_end:%d.%m.%Y}.'))
        if range_end < range_start:
            out.append(_f('error', 'Dates: Range order', 'Range end is before Range start.'))
    pairs = [('dp', ['dp_r1', 'dp_r2'], 'dp_sd', 'Delivery Productivity'),
             ('dss', ['dss_r'], 'dss_sd', 'DSS Usage'), ('cod', ['cod_r'], 'cod_sd', 'COD Digital Transaction'),
             ('lb', ['lb_r'], 'lb_sd', 'Letter Box Clearance')]
    for cat, rkeys, skey, label in pairs:
        sd = frames.get(skey)
        parts = [frames[k] for k in rkeys if frames.get(k) is not None]
        if sd is None or not parts:
            continue
        rng = _totals_by_office(pd.concat(parts, ignore_index=True), cat)
        sdt = _totals_by_office(sd, cat)
        j = pd.concat([rng.rename('r'), sdt.rename('s')], axis=1).dropna()
        over = j[j['s'] > j['r']]
        if len(over):
            out.append(_f('warning', f'{label}: Single Date vs Range',
                          f'{len(over)} office(s) show a larger figure on the Single Date than across the whole Range '
                          f'(e.g. ID {_preview(over.index.astype("int64").tolist(), 4)}). The Range may not cover that date, '
                          f'or the Range and Single Date files may be swapped / in the wrong slot.'))
        elif range_days and range_days > 1 and len(j) and (j['s'] == j['r']).all() and j['r'].sum() > 0:
            out.append(_f('warning', f'{label}: Single Date identical to Range',
                          f'Every office has exactly the same figure on the Single Date as across the {range_days}-day Range. '
                          f'The same file may have been uploaded in both slots.'))
        else:
            out.append(_f('pass', f'{label}: Single Date vs Range', 'Single Date figures are consistent with the Range'))
    return out


# ------------------------------------------------------------- helpers ---

def safe_div(num, den):
    if den is None or pd.isna(den) or den == 0:
        return np.nan
    return num / den


def _ratio(num, den):
    """Vectorised safe_div: NaN where the denominator is missing or zero."""
    num = pd.to_numeric(num, errors='coerce')
    den = pd.to_numeric(den, errors='coerce')
    with np.errstate(divide='ignore', invalid='ignore'):
        r = num / den
    return r.where(den.notna() & (den != 0))


def _num_cols(df, cols):
    d = df.copy()
    for c in cols:
        d[c] = pd.to_numeric(d[c], errors='coerce')
    return d


def _drop_exact_dupes(df):
    """Drop rows that are identical in every column. Used only for the files that carry ONE row per
    office (Delivery Productivity, DSS, Single-Date Letter Box), where an identical repeated line can only
    be an export glitch. Never applied to COD or the per-day Letter Box range file, whose rows can
    legitimately repeat."""
    return df.drop_duplicates()


# --------------------------------------------------------- computation ---

def compute_delivery(master, dp_r1, dp_r2, dp_sd):
    """dp_r1 / dp_r2 / dp_sd are DataFrames or None. Range-1 and Range-2 are the
    two explicit uploads named in the spec (Range-2 exists because the source
    system caps a single export at ~15 days) — summed together per office."""
    out = master[['office_id']].copy()
    for c in ('invoice_r', 'deposit_r', 'invoice_sd', 'deposit_sd'):
        out[c] = np.nan

    cols = ['office-id', 'invoice-count', 'deposit-count']
    parts = [_num_cols(_drop_exact_dupes(d)[cols], cols) for d in (dp_r1, dp_r2) if d is not None]
    if parts:
        dp_r = pd.concat(parts, ignore_index=True)
        g = dp_r.groupby('office-id').agg(invoice_r=('invoice-count', 'sum'),
                                           deposit_r=('deposit-count', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        out = out.drop(columns=['invoice_r', 'deposit_r']).merge(g, on='office_id', how='left')

    if dp_sd is not None:
        d = _num_cols(_drop_exact_dupes(dp_sd)[cols], cols)
        g = d.groupby('office-id').agg(invoice_sd=('invoice-count', 'sum'),
                                        deposit_sd=('deposit-count', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        keep = [c for c in out.columns if c not in ('invoice_sd', 'deposit_sd')]
        out = out[keep].merge(g, on='office_id', how='left')

    out['delivery_pct_r'] = _ratio(out['invoice_r'] - out['deposit_r'], out['invoice_r'])
    out['delivery_pct_sd'] = _ratio(out['invoice_sd'] - out['deposit_sd'], out['invoice_sd'])
    return out[['office_id', 'invoice_r', 'deposit_r', 'invoice_sd', 'deposit_sd',
                'delivery_pct_r', 'delivery_pct_sd']]


def _dss_one(master, dss):
    """One DSS file -> (per-office frame: office_id, pdm, dssart, pct ; excluded-rows frame).
    Offices appearing on several rows are summed, so a duplicated row can never
    duplicate an office in the report."""
    master_ids = set(master['office_id'])
    d = _num_cols(_drop_exact_dupes(dss), ['office_id', 'total_pdm_art_count', 'total_dss_art_count'])
    excl = d[(d['office_id'] == 0) | (~d['office_id'].isin(master_ids))].copy()
    f = d[d['office_id'].isin(master_ids)]
    g = f.groupby('office_id').agg(pdm=('total_pdm_art_count', 'sum'),
                                   dssart=('total_dss_art_count', 'sum')).reset_index()
    g['pct'] = _ratio(g['dssart'], g['pdm'])
    return g, excl


def compute_dss(master, dss_r, dss_sd):
    excl_r = pd.DataFrame(columns=['office_id', 'office_name'])
    excl_sd = pd.DataFrame(columns=['office_id', 'office_name'])
    out = master[['office_id']].copy()
    for c in ('pdm_r', 'dssart_r', 'dss_pct_r', 'pdm_sd', 'dssart_sd', 'dss_pct_sd'):
        out[c] = np.nan

    if dss_r is not None:
        g, excl_r = _dss_one(master, dss_r)
        g = g.rename(columns={'pdm': 'pdm_r', 'dssart': 'dssart_r', 'pct': 'dss_pct_r'})
        out = out.drop(columns=['pdm_r', 'dssart_r', 'dss_pct_r']).merge(g, on='office_id', how='left')
    if dss_sd is not None:
        g, excl_sd = _dss_one(master, dss_sd)
        g = g.rename(columns={'pdm': 'pdm_sd', 'dssart': 'dssart_sd', 'pct': 'dss_pct_sd'})
        out = out.drop(columns=['pdm_sd', 'dssart_sd', 'dss_pct_sd']).merge(g, on='office_id', how='left')
    return out, excl_r, excl_sd


def compute_cod(master, cod_r, cod_sd):
    out = master[['office_id']].copy()
    for c in ('digital_r', 'total_r', 'cod_pct_r', 'digital_sd', 'total_sd', 'cod_pct_sd'):
        out[c] = np.nan
    cols = ['Office ID', 'COD Digital Count', 'Total COD Count']

    if cod_r is not None:
        d = _num_cols(cod_r[cols], cols)
        g = d.groupby('Office ID').agg(digital_r=('COD Digital Count', 'sum'),
                                        total_r=('Total COD Count', 'sum')).reset_index()
        g = g.rename(columns={'Office ID': 'office_id'})
        g['cod_pct_r'] = _ratio(g['digital_r'], g['total_r'])
        out = out.drop(columns=['digital_r', 'total_r', 'cod_pct_r']).merge(g, on='office_id', how='left')

    if cod_sd is not None:
        d = _num_cols(cod_sd[cols], cols)
        g = d.groupby('Office ID').agg(digital_sd=('COD Digital Count', 'sum'),
                                        total_sd=('Total COD Count', 'sum')).reset_index()
        g = g.rename(columns={'Office ID': 'office_id'})
        g['cod_pct_sd'] = _ratio(g['digital_sd'], g['total_sd'])
        out = out.drop(columns=['digital_sd', 'total_sd', 'cod_pct_sd']).merge(g, on='office_id', how='left')

    return out[['office_id', 'digital_r', 'total_r', 'digital_sd', 'total_sd', 'cod_pct_r', 'cod_pct_sd']]


def compute_lb(master, lb_r, lb_sd):
    out = master[['office_id']].copy()
    out['total_letterboxes'] = np.nan
    for c in ('cleared_r', 'boxes_r', 'lb_pct_r', 'cleared_sd', 'boxes_sd', 'lb_pct_sd'):
        out[c] = np.nan
    cols = ['office-id', 'total-cleared', 'total-letterboxes']

    boxes_const_r = None
    if lb_r is not None:
        d = _num_cols(lb_r[cols], cols)
        g = d.groupby('office-id').agg(cleared_r=('total-cleared', 'sum'),
                                        boxes_r=('total-letterboxes', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        g['lb_pct_r'] = _ratio(g['cleared_r'], g['boxes_r'])
        out = out.drop(columns=['cleared_r', 'boxes_r', 'lb_pct_r']).merge(g, on='office_id', how='left')
        boxes_const_r = d.groupby('office-id')['total-letterboxes'].first().reset_index()
        boxes_const_r = boxes_const_r.rename(columns={'office-id': 'office_id', 'total-letterboxes': 'boxes_const_r'})

    if lb_sd is not None:
        d = _num_cols(_drop_exact_dupes(lb_sd)[cols], cols)
        d = d[d['office-id'] != 0]
        g = d.groupby('office-id').agg(cleared_sd=('total-cleared', 'sum'),
                                        boxes_sd=('total-letterboxes', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        g['lb_pct_sd'] = _ratio(g['cleared_sd'], g['boxes_sd'])
        out = out.drop(columns=['cleared_sd', 'boxes_sd', 'lb_pct_sd']).merge(g, on='office_id', how='left')

    if boxes_const_r is not None:
        out = out.merge(boxes_const_r, on='office_id', how='left')
        out['total_letterboxes'] = out['boxes_sd'].where(out['boxes_sd'].notna(), out['boxes_const_r'])
        out = out.drop(columns=['boxes_const_r'])
    else:
        out['total_letterboxes'] = out['boxes_sd']

    return out[['office_id', 'total_letterboxes', 'cleared_r', 'boxes_r', 'cleared_sd', 'boxes_sd',
                'lb_pct_r', 'lb_pct_sd']]


def build_full(master, dp_r1, dp_r2, dp_sd, dss_r, dss_sd, cod_r, cod_sd, lb_r, lb_sd):
    if master['office_id'].duplicated().any():
        raise ValueError('Master has duplicate Office IDs — refusing to build (figures would be double-counted).')
    delivery = compute_delivery(master, dp_r1, dp_r2, dp_sd)
    dss, excl_r, excl_sd = compute_dss(master, dss_r, dss_sd)
    cod = compute_cod(master, cod_r, cod_sd)
    lb = compute_lb(master, lb_r, lb_sd)

    df = master.merge(delivery, on='office_id', how='left')
    df = df.merge(dss, on='office_id', how='left')
    df = df.merge(cod, on='office_id', how='left')
    df = df.merge(lb, on='office_id', how='left')
    assert len(df) == len(master), 'internal error: join changed the number of offices'

    df.loc[df['office_type'] == 'BPO', ['lb_pct_r', 'lb_pct_sd', 'total_letterboxes']] = np.nan

    avail = {
        'delivery_r': (dp_r1 is not None) or (dp_r2 is not None), 'delivery_sd': dp_sd is not None,
        'dss_r': dss_r is not None, 'dss_sd': dss_sd is not None,
        'cod_r': cod_r is not None, 'cod_sd': cod_sd is not None,
        'lb_r': lb_r is not None, 'lb_sd': lb_sd is not None,
    }
    return df, excl_r, excl_sd, avail


def check_percentage_bounds(df):
    """Spec §8.2 — every percentage must be in [0%, 100%]. Returns findings."""
    cols = ['delivery_pct_r', 'delivery_pct_sd', 'dss_pct_r', 'dss_pct_sd',
            'cod_pct_r', 'cod_pct_sd', 'lb_pct_r', 'lb_pct_sd']
    bad = []
    for c in cols:
        v = df[c]
        n = int(((v < 0) | (v > 1)).sum())
        if n:
            bad.append(f'{c}: {n} office(s)')
    if bad:
        return [_f('warning', 'Percentage bounds (0–100%)', '; '.join(bad))]
    return [_f('pass', 'Percentage bounds (0–100%)', 'All percentages are within 0–100%')]


# --------------------------------------------------------- aggregation ---

def weighted_row(df):
    return {
        'delivery_pct_r': safe_div((df['invoice_r'] - df['deposit_r']).sum(skipna=True), df['invoice_r'].sum(skipna=True)),
        'delivery_pct_sd': safe_div((df['invoice_sd'] - df['deposit_sd']).sum(skipna=True), df['invoice_sd'].sum(skipna=True)),
        'dss_pct_r': safe_div(df['dssart_r'].sum(skipna=True), df['pdm_r'].sum(skipna=True)),
        'dss_pct_sd': safe_div(df['dssart_sd'].sum(skipna=True), df['pdm_sd'].sum(skipna=True)),
        'cod_pct_r': safe_div(df['digital_r'].sum(skipna=True), df['total_r'].sum(skipna=True)),
        'cod_pct_sd': safe_div(df['digital_sd'].sum(skipna=True), df['total_sd'].sum(skipna=True)),
        'lb_pct_r': safe_div(df['cleared_r'].sum(skipna=True), df['boxes_r'].sum(skipna=True)),
        'lb_pct_sd': safe_div(df['cleared_sd'].sum(skipna=True), df['boxes_sd'].sum(skipna=True)),
    }


def build_divisional_summary(df, total_label='Division Total'):
    """Sub-division row order: alphabetical by name (derived from the Master),
    Division Total always last. Each row carries an explicit `is_total` flag."""
    tables = {}
    for key, filt in [('excl_bo', df['office_type'] != 'BPO'),
                       ('only_bo', df['office_type'] == 'BPO'),
                       ('incl_bo', df['office_type'].notna())]:
        sub = df[filt]
        rows = []
        for sd in get_sub_divisions(df):
            row = weighted_row(sub[sub['sub_division'] == sd])
            row['sub_division'] = sd
            row['is_total'] = False
            rows.append(row)
        total_row = weighted_row(sub)
        total_row['sub_division'] = total_label
        total_row['is_total'] = True
        rows.append(total_row)
        tables[key] = rows
    return tables


def build_excl_bo_table(df):
    """All non-BPO (SPO+HPO) offices. Sort rule: offices with a genuine
    Delivery % (Range) value first, sorted by Delivery % (Range) ASCENDING
    (worst performer first); non-delivery offices (blank Delivery % R)
    unchanged — still grouped at the bottom, A-Z by name."""
    sub = df[df['office_type'] != 'BPO'].copy()
    has_data = sub[sub['delivery_pct_r'].notna()].sort_values(
        ['delivery_pct_r', 'office_name'], ascending=[True, True])
    no_data = sub[sub['delivery_pct_r'].isna()].sort_values('office_name')
    return pd.concat([has_data, no_data], ignore_index=False)


def build_defaulters(df):
    """Each of the 3 tables (Delivery/DSS/COD) sorted by that metric's own
    Range % ascending — worst defaulter first. One entry per sub-division."""
    out = {}
    for sd in get_sub_divisions(df):
        s = df[df['sub_division'] == sd]
        out[sd] = {
            'delivery': s[s['delivery_pct_r'] < 0.90].sort_values(
                ['delivery_pct_r', 'office_name'], ascending=[True, True]),
            'dss': s[s['dss_pct_r'] < 0.90].sort_values(
                ['dss_pct_r', 'office_name'], ascending=[True, True]),
            'cod': s[s['cod_pct_r'] < 0.90].sort_values(
                ['cod_pct_r', 'office_name'], ascending=[True, True]),
        }
    return out


def build_exceptional_defaulters(df):
    params = [
        ('Delivery % (Range)', 'delivery_pct_r'), ('Delivery % (Single Date)', 'delivery_pct_sd'),
        ('DSS % (Range)', 'dss_pct_r'), ('DSS % (Single Date)', 'dss_pct_sd'),
        ('COD % (Range)', 'cod_pct_r'), ('COD % (Single Date)', 'cod_pct_sd'),
        ('LB % (Range)', 'lb_pct_r'), ('LB % (Single Date)', 'lb_pct_sd'),
    ]
    rows = []
    for _, r in df.iterrows():
        for label, col in params:
            v = r[col]
            if pd.notna(v) and v < 0.30:
                rows.append({'office_id': r['office_id'], 'office_name': r['office_name'],
                              'sub_division': r['sub_division'], 'office_type': r['office_type'],
                              'parameter': label, 'value': v})
    result = pd.DataFrame(rows)
    if len(result):
        result = result.sort_values(['sub_division', 'office_name', 'parameter']).reset_index(drop=True)
    return result
