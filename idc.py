"""
IDC (Instant Delivery Centre) Delivery report — engine + workbook builder.

Inputs (two CSV exports, plus calendar dates chosen in the UI — never read from filenames):
  * IDC DPR  — Date, Office Name, Product, Beat, Postman, Invoiced, Delivered, Returned, Deposited, Redirected ...
  * IDC DSS  — Office Name, Invoiced Date, Beat Name, Postman Name, PDMS Invoiced Count, DSS Invoiced Count ...

One row per delivery person (postman), sorted by Delivery % (Range) — smallest first.

KPIs (all volume-weighted: sum / sum; blank, never 0, when the denominator is 0)
  Delivery %          = (Invoiced - Deposited) / Invoiced            all products
  DSS Usage %         = DSS Invoiced Count / PDMS Invoiced Count
  Parcel Delivery %   = (Invoiced - Deposited) / Invoiced            products whose name contains "parcel"
  RTS %               = Returned / Invoiced
  Redirected %        = Redirected / Invoiced
Delivery % and DSS Usage % are shown for the Range and for the Single Date.
"""
import datetime
import io
import re

import numpy as np
import pandas as pd
import xlsxwriter

DPR_REQUIRED = ['Date', 'Office Name', 'Product', 'Postman', 'Invoiced', 'Delivered', 'Returned',
                'Deposited', 'Redirected']
DSS_REQUIRED = ['Office Name', 'Invoiced Date', 'Postman Name', 'PDMS Invoiced Count', 'DSS Invoiced Count']
DPR_NUM = ['Invoiced', 'Delivered', 'Returned', 'Deposited', 'Redirected']
DSS_NUM = ['PDMS Invoiced Count', 'DSS Invoiced Count']

# thresholds (percent of 1.0)
DELIVERY_RED, DELIVERY_GREEN = 0.93, 0.95
LIMIT = 0.05            # RTS % / Redirected % flagged above this

INSTRUCTION_RUNS = [
    'सर्व डिलीवरी पोस्टमन यांनी चालू महिन्याच्या ०१ तारखेपासून कालपर्यंत आणि कालचा  डिलीवरी%, DSS %, '
    'एकूण महिन्याचा पार्सल डिलीवरी % पाहावे. ',
    '. सहायक अधीक्षक  तसेच मेल PA, मेल ओवरसीयर, ऑफिस इन्चार्ज PM',
    ' यांनी नियमित Supervision ठेऊन पुढील दिवसांमध्ये कमीत कमी आर्टिकल deposit राहतील याची दक्षता घ्यावी. '
    'रिमार्क व्हेरिफिकेशन रजिस्टर नियमित अपडेट ठेवावे.',
]
from signature import default_signature, clean_lines  # noqa: E402


# ------------------------------------------------------------------ helpers ---
def person_key(s):
    """'MADHAV  DAWARE' -> 'Madhav Daware' (collapse spaces, Proper Case) so both files join on one name."""
    return re.sub(r'\s+', ' ', str(s)).strip().title()


def office_short(name):
    """'Karad H.O' -> 'Karad HO' (used in the report title)."""
    return re.sub(r'\s+', ' ', str(name).replace('.', '')).strip()


def range_labels(r0, r1):
    """(title_label, header_label): '01- 13 October' / '01-13 October'; spans months -> '01 Sep - 13 Oct'."""
    if (r0.year, r0.month) == (r1.year, r1.month):
        return f"{r0.day:02d}- {r1.day:02d} {r1:%B}", f"{r0.day:02d}-{r1.day:02d} {r1:%B}"
    s = f"{r0.day:02d} {r0:%b} - {r1.day:02d} {r1:%b}"
    return s, s


def _norm_cols(df):
    """Match column names ignoring case / extra spaces; returns df with canonical spellings."""
    want = {c.lower(): c for c in DPR_REQUIRED + DSS_REQUIRED + ['Beat', 'Beat Name', 'Batch']}
    return df.rename(columns={c: want.get(re.sub(r'\s+', ' ', str(c)).strip().lower(), c) for c in df.columns})


def _parse_dates(series):
    return pd.to_datetime(series.astype(str).str.strip(), format='%d-%m-%Y', errors='coerce')


def detect_kind(df):
    cols = set(df.columns)
    if {'PDMS Invoiced Count', 'DSS Invoiced Count'} <= cols:
        return 'dss'
    if {'Invoiced', 'Deposited', 'Product'} <= cols:
        return 'dpr'
    if {'Office ID', 'Office Name'} <= cols and 'Invoiced' not in cols:
        return 'other'
    return 'unknown'


def prepare(dpr, dss):
    """Normalise names/dates/numbers. Returns (dpr, dss) prepared copies; either may be None."""
    out = []
    for df, nums, dcol, pcol in ((dpr, DPR_NUM, 'Date', 'Postman'), (dss, DSS_NUM, 'Invoiced Date', 'Postman Name')):
        if df is None:
            out.append(None)
            continue
        d = _norm_cols(df).copy()
        d['Person'] = d[pcol].map(person_key)
        d['D'] = _parse_dates(d[dcol]).dt.date
        for c in nums:
            d[c] = pd.to_numeric(d[c], errors='coerce')
        out.append(d)
    dpr_p, dss_p = out
    if dpr_p is not None:
        dpr_p['Parcel'] = dpr_p['Product'].astype(str).str.contains('parcel', case=False)
    return dpr_p, dss_p


def offices_in(dpr, dss):
    names = set()
    for d in (dpr, dss):
        if d is not None and 'Office Name' in d.columns:
            names |= set(d['Office Name'].dropna().astype(str).str.strip())
    return sorted(names)


def filter_office(df, office):
    if df is None or office is None:
        return df
    return df[df['Office Name'].astype(str).str.strip() == office]


# -------------------------------------------------------------- validation ---
def _f(level, check, detail):
    return {'level': level, 'check': check, 'detail': detail}


def run_validations(dpr_raw, dss_raw, single_date, r0, r1):
    """Findings: list of {'level': pass|info|warning|error, 'check', 'detail'}. Errors block the report."""
    F = []
    if dpr_raw is None:
        return [_f('error', 'IDC DPR file', 'Upload the IDC DPR (delivery performance) CSV — it is required.')]
    dpr_raw = _norm_cols(dpr_raw)
    kind = detect_kind(dpr_raw)
    if kind == 'dss':
        F.append(_f('error', 'IDC DPR file', 'This looks like the IDC DSS file (it has PDMS / DSS Invoiced Count). '
                                              'Upload it in the DSS slot and the DPR file here.'))
    missing = [c for c in DPR_REQUIRED if c not in dpr_raw.columns]
    if missing:
        F.append(_f('error', 'IDC DPR file', f"Missing column(s): {', '.join(missing)}."))
    if dss_raw is not None:
        dss_raw = _norm_cols(dss_raw)
        if detect_kind(dss_raw) == 'dpr':
            F.append(_f('error', 'IDC DSS file', 'This looks like the IDC DPR file (it has Product / Deposited). '
                                                  'Upload it in the DPR slot and the DSS file here.'))
        miss = [c for c in DSS_REQUIRED if c not in dss_raw.columns]
        if miss:
            F.append(_f('error', 'IDC DSS file', f"Missing column(s): {', '.join(miss)}."))
    if r0 is None or r1 is None or single_date is None:
        F.append(_f('error', 'Dates', 'Pick the Range From / To dates and the Single Date.'))
        return F
    if r1 < r0:
        F.append(_f('error', 'Dates', 'Range end is before Range start.'))
    if not (r0 <= single_date <= r1):
        F.append(_f('warning', 'Dates', f'Single Date {single_date:%d-%m-%Y} is outside the Range '
                                          f'{r0:%d-%m-%Y} to {r1:%d-%m-%Y}.'))
    if any(x['level'] == 'error' for x in F):
        return F

    dpr, dss = prepare(dpr_raw, dss_raw)
    for name, d, nums, dcol in (('IDC DPR file', dpr, DPR_NUM, 'Date'), ('IDC DSS file', dss, DSS_NUM, 'Invoiced Date')):
        if d is None:
            continue
        if len(d) == 0:
            F.append(_f('error', name, 'The file has no data rows.'))
            continue
        for c in nums:
            bad = d[c].isna().sum()
            if bad:
                F.append(_f('error', name, f"Column '{c}' has {bad} blank or non-numeric value(s)."))
            elif (d[c] < 0).any():
                F.append(_f('error', name, f"Column '{c}' has negative values."))
        nbad = d['D'].isna().sum()
        if nbad:
            F.append(_f('error', name, f"{nbad} row(s) have a date in '{dcol}' that is not dd-mm-yyyy."))
    if any(x['level'] == 'error' for x in F):
        return F

    # offices
    offs = offices_in(dpr, dss)
    if len(offs) > 1:
        F.append(_f('info', 'Offices', f"Files contain {len(offs)} offices ({', '.join(offs)}); choose one in the "
                                         "'Office' selector — one report is built per office."))
    for name, d in (('IDC DPR file', dpr), ('IDC DSS file', dss)):
        if d is None:
            continue
        outside = d[(d['D'] < r0) | (d['D'] > r1)]
        if len(outside):
            F.append(_f('warning', name, f"{len(outside)} row(s) are dated outside the Range "
                                          f"({outside['D'].min():%d-%m-%Y} … {outside['D'].max():%d-%m-%Y}); "
                                          "they are still counted in the Range columns."))
        if not (d['D'] == single_date).any():
            F.append(_f('warning', name, f"No rows dated {single_date:%d-%m-%Y} — the Single Date columns will be blank."))
        dup = int(d.duplicated(subset=[c for c in d.columns if c not in ('Person', 'D', 'Parcel')]).sum())
        if dup:
            F.append(_f('warning', name, f"{dup} exact repeated row(s) found. They are counted as exported "
                                          "(not removed) — check the export."))
    # integrity
    parts = dpr['Delivered'] + dpr['Returned'] + dpr['Deposited'] + dpr['Redirected']
    if 'Returns Not Taken' in dpr.columns:
        parts = parts + pd.to_numeric(dpr['Returns Not Taken'], errors='coerce').fillna(0)
    nbal = int((dpr['Invoiced'] != parts).sum())
    if nbal:
        F.append(_f('warning', 'IDC DPR file', f"{nbal} row(s) where Invoiced ≠ Delivered + Returned + Deposited + "
                                                "Redirected (+ Returns Not Taken)."))
    else:
        F.append(_f('pass', 'IDC DPR file', 'Every row balances: Invoiced = Delivered + Returned + Deposited + Redirected.'))
    if (dpr['Deposited'] > dpr['Invoiced']).any():
        F.append(_f('warning', 'IDC DPR file', 'Some rows have Deposited > Invoiced.'))
    if dss is not None:
        if (dss['DSS Invoiced Count'] > dss['PDMS Invoiced Count']).any():
            F.append(_f('warning', 'IDC DSS file', 'Some rows have DSS Invoiced Count > PDMS Invoiced Count.'))
        a, b = set(dpr['Person']), set(dss['Person'])
        if a - b:
            F.append(_f('warning', 'Delivery persons', f"In DPR but not in DSS ({len(a - b)}): {', '.join(sorted(a - b))}. "
                                                         "Their DSS cells will be blank."))
        if b - a:
            F.append(_f('warning', 'Delivery persons', f"In DSS but not in DPR ({len(b - a)}): {', '.join(sorted(b - a))}. "
                                                         "Their delivery cells will be blank."))
        if not (a - b) and not (b - a):
            F.append(_f('pass', 'Delivery persons', f'All {len(a)} delivery persons appear in both files.'))
    else:
        F.append(_f('warning', 'IDC DSS file', 'Not provided — DSS Usage % columns will be blank.'))
    # missing calendar days inside the range
    have = set(dpr['D']) | (set(dss['D']) if dss is not None else set())
    gaps = [r0 + datetime.timedelta(days=i) for i in range((r1 - r0).days + 1)
            if (r0 + datetime.timedelta(days=i)) not in have]
    if gaps:
        F.append(_f('info', 'Dates', f"No data on {len(gaps)} day(s) of the Range: "
                                      f"{', '.join(g.strftime('%d-%m') for g in gaps)} (holidays/Sundays are normal)."))
    return F


# ------------------------------------------------------------------ compute ---
def _pct(n, d):
    return np.nan if (d is None or d == 0 or pd.isna(d)) else n / d


def compute_table(dpr, dss, single_date):
    """Return (table DataFrame sorted for the report, totals dict). dpr/dss are *prepared* frames (dss may be None)."""
    persons = set(dpr['Person']) | (set(dss['Person']) if dss is not None else set())
    rows = []
    for p in persons:
        a = dpr[dpr.Person == p]
        a1 = a[a.D == single_date]
        ap = a[a.Parcel]
        inv = a.Invoiced.sum()
        row = dict(person=p, inv=inv,
                   dl_r=_pct(inv - a.Deposited.sum(), inv),
                   dl_s=_pct(a1.Invoiced.sum() - a1.Deposited.sum(), a1.Invoiced.sum()),
                   p_inv=ap.Invoiced.sum(), p_dl=_pct(ap.Invoiced.sum() - ap.Deposited.sum(), ap.Invoiced.sum()),
                   rts=_pct(a.Returned.sum(), inv), red=_pct(a.Redirected.sum(), inv),
                   ds_r=np.nan, ds_s=np.nan)
        if dss is not None:
            s = dss[dss.Person == p]
            s1 = s[s.D == single_date]
            row['ds_r'] = _pct(s['DSS Invoiced Count'].sum(), s['PDMS Invoiced Count'].sum())
            row['ds_s'] = _pct(s1['DSS Invoiced Count'].sum(), s1['PDMS Invoiced Count'].sum())
        rows.append(row)
    T = pd.DataFrame(rows)
    T['_k'] = T.dl_r.fillna(9)                       # no-invoice rows go last
    T = T.sort_values(['_k', 'person']).drop(columns='_k').reset_index(drop=True)
    tinv = dpr.Invoiced.sum()
    t1 = dpr[dpr.D == single_date]
    tp = dpr[dpr.Parcel]
    TOT = dict(dl_r=_pct(tinv - dpr.Deposited.sum(), tinv),
               dl_s=_pct(t1.Invoiced.sum() - t1.Deposited.sum(), t1.Invoiced.sum()),
               p_inv=tp.Invoiced.sum(), p_dl=_pct(tp.Invoiced.sum() - tp.Deposited.sum(), tp.Invoiced.sum()),
               rts=_pct(dpr.Returned.sum(), tinv), red=_pct(dpr.Redirected.sum(), tinv),
               ds_r=np.nan, ds_s=np.nan)
    if dss is not None:
        TOT['ds_r'] = _pct(dss['DSS Invoiced Count'].sum(), dss['PDMS Invoiced Count'].sum())
        s1 = dss[dss.D == single_date]
        TOT['ds_s'] = _pct(s1['DSS Invoiced Count'].sum(), s1['PDMS Invoiced Count'].sum())
    return T, TOT


# ----------------------------------------------------------------- workbook ---
def write_idc_sheets(wb, dpr, dss, office_label, record_prefix, single_date, r0, r1, findings=None,
                     signature=None, standalone=True):
    """Write the 'IDC Report' sheet (+ hidden Src_IDC_DPR / Src_IDC_DSS) into an existing xlsxwriter workbook.
    dpr/dss are prepared frames already filtered to one office. standalone=True also adds the hidden
    'Validation Log' (inside the MMU workbook the findings go into that workbook's own log instead)."""
    T, TOT = compute_table(dpr, dss, single_date)
    n = len(T)
    title_lbl, head_lbl = range_labels(r0, r1)
    have_dss = dss is not None

    A = 'Arial'
    MED, THN = 2, 1
    BLUE, PINK = '#8DB4E2', '#E6B8B7'

    def cellfmt(col, role, **extra):
        left = MED if col in (2, 4, 6, 8) else THN
        right = MED if col in (3, 5, 7, 9) else THN
        top = bot = THN
        if role == 'h4':
            top, bot = MED, (THN if col < 2 else MED); left = MED if col in (0, 2, 4, 6, 8) else THN
        if role == 'h5':
            top, bot = 0, (THN if col < 2 else MED); left = MED if col in (0, 2, 4, 6, 8) else THN
        if role == 'last':
            bot = 0
        if role == 'total':
            top, bot = MED, MED; left = MED if col in (0, 2, 4, 6, 8) else THN
        if role == 'ins':
            top, bot = THN, 0; left = right = THN
        d = dict(font_name=A, font_size=12, align='center', valign='vcenter', text_wrap=True,
                 left=left, right=right, top=top, bottom=bot)
        if role in ('h4', 'h5'):
            d['bold'] = True
            d['bg_color'] = PINK if (role == 'h4' and col in (4, 8)) else BLUE
        if role in ('ins', 'total'):
            d['bold'] = True
        if role == 'total':
            d['text_wrap'] = False
        d.update(extra)
        return wb.add_format(d)

    ws = wb.add_worksheet('IDC Report')
    sp = wb.add_worksheet('Src_IDC_DPR'); sp.hide()
    sd = wb.add_worksheet('Src_IDC_DSS'); sd.hide()
    vl = None
    if standalone:
        vl = wb.add_worksheet('Validation Log'); vl.hide()

    # ---- hidden source sheets (the report's SUMIFS read these)
    dfmt = wb.add_format({'num_format': 'dd-mm-yyyy'})
    for j, h in enumerate(['Date', 'Person', 'Product', 'Invoiced', 'Returned', 'Deposited', 'Redirected', 'IsParcel']):
        sp.write(0, j, h)
    for i, r in enumerate(dpr.itertuples(), 1):
        sp.write_datetime(i, 0, datetime.datetime.combine(r.D, datetime.time()), dfmt)
        sp.write(i, 1, r.Person); sp.write(i, 2, r.Product)
        sp.write(i, 3, r.Invoiced); sp.write(i, 4, r.Returned); sp.write(i, 5, r.Deposited)
        sp.write(i, 6, r.Redirected); sp.write(i, 7, int(r.Parcel))
    for j, h in enumerate(['Date', 'Person', 'PDMS Invoiced', 'DSS Invoiced']):
        sd.write(0, j, h)
    if have_dss:
        for i, (d_, p_, a_, b_) in enumerate(zip(dss['D'], dss['Person'], dss['PDMS Invoiced Count'],
                                                  dss['DSS Invoiced Count']), 1):
            sd.write_datetime(i, 0, datetime.datetime.combine(d_, datetime.time()), dfmt)
            sd.write(i, 1, p_); sd.write(i, 2, a_); sd.write(i, 3, b_)

    # ---- validation log (hidden)
    if vl is not None:
        for j, h in enumerate(['Check', 'Detail', 'Result']):
            vl.write(0, j, h)
        lvl = {'pass': 'PASS', 'info': 'INFO', 'warning': 'WARNING', 'error': 'FAIL'}
        rows = [('Report', f"IDC {office_label}; record prefix {record_prefix}", 'INFO'),
                ('Dates (from calendar pickers, never filenames)',
                 f"Single Date={single_date}; Range={r0} to {r1}", 'INFO')]
        rows += [(x['check'], x['detail'], lvl.get(x['level'], x['level'].upper())) for x in (findings or [])]
        for i, (a, b, c) in enumerate(rows, 1):
            vl.write(i, 0, a); vl.write(i, 1, b); vl.write(i, 2, c)
        vl.set_column(0, 0, 40); vl.set_column(1, 1, 110); vl.set_column(2, 2, 14)

    # ---- report sheet
    f_title = wb.add_format(dict(font_name=A, font_size=16, bold=True, align='center', valign='vcenter'))
    f_rec = wb.add_format(dict(font_name=A, font_size=12, align='right'))
    f_sig = wb.add_format(dict(font_name=A, font_size=12, bold=True, align='left', valign='vcenter'))
    f_b = wb.add_format(dict(font_name=A, font_size=12, bold=True))
    f_br = wb.add_format(dict(font_name=A, font_size=12, bold=True, font_color='#FF0000'))
    ws.set_column('A:A', 6.5); ws.set_column('B:B', 31); ws.set_column('C:J', 12.5)
    ws.merge_range('A1:J1', f"IDC {office_label} Delivery % ({title_lbl})", f_title); ws.set_row(0, 30)
    ws.merge_range('A2:J2', f"{record_prefix}/G/MMU/IDC/Delivery Performance/{single_date:%d.%m.%Y}", f_rec)
    ws.merge_range('A3:J3', '', cellfmt(0, 'ins'))
    ws.write_rich_string('A3', f_b, INSTRUCTION_RUNS[0], f_br, INSTRUCTION_RUNS[1], f_b, INSTRUCTION_RUNS[2],
                         cellfmt(0, 'ins'))
    ws.set_row(2, 92)
    H4 = [cellfmt(c, 'h4') for c in range(10)]
    H5 = [cellfmt(c, 'h5') for c in range(10)]
    H5d = [cellfmt(c, 'h5', num_format='d-mmm') for c in range(10)]
    ws.merge_range('A4:A5', 'Sr. No', H4[0]); ws.merge_range('B4:B5', 'Delivery Person', H4[1])
    ws.merge_range('C4:D4', 'Delivery %', H4[2]); ws.merge_range('E4:F4', 'DSS Usage %', H4[4])
    ws.merge_range('G4:H4', f'Parcels ({head_lbl})', H4[6]); ws.merge_range('I4:J4', head_lbl, H4[8])
    sdt = datetime.datetime.combine(single_date, datetime.time())
    ws.write('C5', f'({head_lbl})', H5[2]); ws.write_datetime('D5', sdt, H5d[3])
    ws.write('E5', f'({head_lbl})', H5[4]); ws.write_datetime('F5', sdt, H5d[5])
    ws.write('G5', 'Invoiced', H5[6]); ws.write('H5', 'Delivery %', H5[7])
    ws.write('I5', 'RTS %', H5[8]); ws.write('J5', 'Redirected %', H5[9])
    ws.set_row(3, 30); ws.set_row(4, 45)

    P, S = 'Src_IDC_DPR!', 'Src_IDC_DSS!'
    v = lambda x: '' if (x is None or (isinstance(x, float) and np.isnan(x))) else x
    r0i = 5
    tr = r0i + n

    def body(c, i, **k):
        return cellfmt(c, 'last' if i == n - 1 else 'body', **k)

    for i, r in T.iterrows():
        rr = r0i + i; b = f'$B{rr + 1}'
        pc = lambda c: body(c, i, num_format='0.00%', text_wrap=False)
        ws.write(rr, 0, i + 1, body(0, i, text_wrap=False))
        ws.write(rr, 1, r.person, body(1, i, align='left', text_wrap=False))
        inv = f"SUMIFS({P}$D:$D,{P}$B:$B,{b})"; dep = f"SUMIFS({P}$F:$F,{P}$B:$B,{b})"
        ws.write_formula(rr, 2, f'=IF({inv}=0,"",({inv}-{dep})/{inv})', pc(2), v(r.dl_r))
        inv1 = f"SUMIFS({P}$D:$D,{P}$B:$B,{b},{P}$A:$A,$D$5)"; dep1 = f"SUMIFS({P}$F:$F,{P}$B:$B,{b},{P}$A:$A,$D$5)"
        ws.write_formula(rr, 3, f'=IF({inv1}=0,"",({inv1}-{dep1})/{inv1})', pc(3), v(r.dl_s))
        if have_dss:
            pdm = f"SUMIFS({S}$C:$C,{S}$B:$B,{b})"; dsv = f"SUMIFS({S}$D:$D,{S}$B:$B,{b})"
            ws.write_formula(rr, 4, f'=IF({pdm}=0,"",{dsv}/{pdm})', pc(4), v(r.ds_r))
            pdm1 = f"SUMIFS({S}$C:$C,{S}$B:$B,{b},{S}$A:$A,$F$5)"; dsv1 = f"SUMIFS({S}$D:$D,{S}$B:$B,{b},{S}$A:$A,$F$5)"
            ws.write_formula(rr, 5, f'=IF({pdm1}=0,"",{dsv1}/{pdm1})', pc(5), v(r.ds_s))
        else:
            ws.write_blank(rr, 4, None, pc(4)); ws.write_blank(rr, 5, None, pc(5))
        pinv = f"SUMIFS({P}$D:$D,{P}$B:$B,{b},{P}$H:$H,1)"; pdep = f"SUMIFS({P}$F:$F,{P}$B:$B,{b},{P}$H:$H,1)"
        ws.write_formula(rr, 6, f'={pinv}', body(6, i, num_format='0', text_wrap=False), int(r.p_inv))
        ws.write_formula(rr, 7, f'=IF({pinv}=0,"",({pinv}-{pdep})/{pinv})', pc(7), v(r.p_dl))
        ws.write_formula(rr, 8, f'=IF({inv}=0,"",SUMIFS({P}$E:$E,{P}$B:$B,{b})/{inv})', pc(8), v(r.rts))
        ws.write_formula(rr, 9, f'=IF({inv}=0,"",SUMIFS({P}$G:$G,{P}$B:$B,{b})/{inv})', pc(9), v(r.red))

    T_ = lambda c, **k: cellfmt(c, 'total', **k)
    tp = lambda c: T_(c, num_format='0.00%')
    ws.merge_range(tr, 0, tr, 1, f'IDC {office_label} Total', T_(0, align='left'))
    INV, DEP = f"SUM({P}$D:$D)", f"SUM({P}$F:$F)"
    ws.write_formula(tr, 2, f'=IF({INV}=0,"",({INV}-{DEP})/{INV})', tp(2), v(TOT['dl_r']))
    i1 = f"SUMIFS({P}$D:$D,{P}$A:$A,$D$5)"; d1 = f"SUMIFS({P}$F:$F,{P}$A:$A,$D$5)"
    ws.write_formula(tr, 3, f'=IF({i1}=0,"",({i1}-{d1})/{i1})', tp(3), v(TOT['dl_s']))
    if have_dss:
        ws.write_formula(tr, 4, f'=IF(SUM({S}$C:$C)=0,"",SUM({S}$D:$D)/SUM({S}$C:$C))', tp(4), v(TOT['ds_r']))
        p1 = f"SUMIFS({S}$C:$C,{S}$A:$A,$F$5)"
        ws.write_formula(tr, 5, f'=IF({p1}=0,"",SUMIFS({S}$D:$D,{S}$A:$A,$F$5)/{p1})', tp(5), v(TOT['ds_s']))
    else:
        ws.write_blank(tr, 4, None, tp(4)); ws.write_blank(tr, 5, None, tp(5))
    PI, PD = f"SUMIFS({P}$D:$D,{P}$H:$H,1)", f"SUMIFS({P}$F:$F,{P}$H:$H,1)"
    ws.write_formula(tr, 6, f'={PI}', T_(6, num_format='0'), int(TOT['p_inv']))
    ws.write_formula(tr, 7, f'=IF({PI}=0,"",({PI}-{PD})/{PI})', tp(7), v(TOT['p_dl']))
    ws.write_formula(tr, 8, f'=IF({INV}=0,"",SUM({P}$E:$E)/{INV})', tp(8), v(TOT['rts']))
    ws.write_formula(tr, 9, f'=IF({INV}=0,"",SUM({P}$G:$G)/{INV})', tp(9), v(TOT['red']))

    # ---- conditional formatting
    RED, AMB, GRN = '#C00000', '#B8860B', '#00803C'
    fc = lambda c, **k: wb.add_format(dict(font_color=c, **k))
    first, lastd, totx = r0i + 1, tr, tr + 1

    def fx(col, a, b_, cond, fmt):
        ws.conditional_format(f'{col}{a}:{col}{b_}', {'type': 'formula',
                              'criteria': '=' + cond.replace('@', f'{col}{a}'), 'format': fmt})

    for col in 'CDH':                                   # colour scale: Delivery %, Parcel Delivery % (data rows)
        ws.conditional_format(f'{col}{first}:{col}{lastd}', {'type': '3_color_scale', 'min_color': '#F8696B',
                              'mid_color': '#FFEB84', 'max_color': '#63BE7B', 'mid_type': 'percentile', 'mid_value': 50})
    ws.conditional_format(f'G{first}:G{lastd}', {'type': 'data_bar', 'bar_color': '#FFB628',
                          'bar_border_color': '#FFB628', 'data_bar_2010': True, 'bar_solid': False})
    fx('E', first, totx, 'AND(ISNUMBER(@),ROUND(@,4)<1)', fc(RED, bold=True))   # DSS < 100% -> red text
    fx('F', first, totx, 'AND(ISNUMBER(@),ROUND(@,4)<1)', fc(RED))
    for col in 'IJ':                                    # RTS / Redirected above 5%: dark-red text on light-red fill
        fx(col, first, lastd, f'AND(ISNUMBER(@),@>{LIMIT})',
           wb.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE'}))
    for col in 'CDH':                                   # total row: text colours only
        fx(col, totx, totx, f'AND(ISNUMBER(@),ROUND(@,4)<{DELIVERY_RED})', fc(RED))
        fx(col, totx, totx, f'AND(ISNUMBER(@),ROUND(@,4)>={DELIVERY_RED},ROUND(@,4)<{DELIVERY_GREEN})', fc(AMB))
        fx(col, totx, totx, f'AND(ISNUMBER(@),ROUND(@,4)>={DELIVERY_GREEN})', fc(GRN))
    for col in 'IJ':
        fx(col, totx, totx, f'AND(ISNUMBER(@),ROUND(@,4)>={LIMIT})', fc(RED))
        fx(col, totx, totx, f'AND(ISNUMBER(@),ROUND(@,4)<{LIMIT})', fc(GRN))

    sr = tr + 3
    sig = clean_lines(signature if signature is not None else default_signature('Karad'))
    for k, line in enumerate(sig):
        ws.write(sr + k, 4, line, f_sig)
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1); ws.set_margins(0.4, 0.4, 0.5, 0.5)
    ws.hide_gridlines(2); ws.print_area(0, 0, sr + max(len(sig), 1) - 1, 9)
    if standalone:
        ws.activate()
    return T, TOT


def build_workbook(dpr, dss, office_label, record_prefix, single_date, r0, r1, findings=None, signature=None):
    """Stand-alone IDC workbook (used by the IDC page). Returns (BytesIO, table, totals)."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True})
    T, TOT = write_idc_sheets(wb, dpr, dss, office_label, record_prefix, single_date, r0, r1, findings,
                              signature, standalone=True)
    wb.close()
    buf.seek(0)
    return buf, T, TOT
