"""
workbook.py — builds the MMU Hub Report .xlsx (Build Spec v10 layout) via
xlsxwriter, from already-computed engine.py data. Returns an in-memory
BytesIO buffer ready for st.download_button. Division-agnostic: the number
and names of sub-divisions, the division name and the office counts all come
from the data, not from constants.

Raw Data formulas are written WITH cached values (xlsxwriter's
write_formula(..., value=...)) so the file shows correct numbers the
instant it's opened, with no LibreOffice/Excel recalculation pass needed —
Excel will still recalculate live if the underlying hidden sheets are
edited by hand later.
"""
import io
import re
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import timedelta
from signature import default_signature, clean_lines
from engine import (get_sub_divisions, default_record_prefix, weighted_row, build_divisional_summary,
                     build_excl_bo_table, build_defaulters, build_exceptional_defaulters)

NAVY = '#1F4E78'
RED = '#C00000'
ORANGE = '#E67E22'
PURPLE = '#7030A0'
PEACH = '#FCE4D6'
REDTEXT = '#FF0000'
GRID = '#808080'
GROUP_BORDER = '#2C2C2C'

TEXT_A = ("चालू महिन्याच्या ०१ तारखेपासून काल पर्यंत आणि काल दिवशीचा आपल्या सब ऑफिस मधील MMU "
          "संबंधित पर्फोर्मंस पाहावे. सर्व ऑफिस इन्चार्ज PM/SPM हे ठरलेल्या KPI प्रमाणे होण्यासाठी "
          "योग्य ते Supervision ठेवतील.")
TEXT_B1 = ("या यादीतील ऑफिसेस मध्ये चालू महिन्याच्या ०१ तारखेपासून कालपर्यंत डिलीवरी% हा ९०% पेक्षा "
           "कमी आहे. ऑफिस चे इन्चार्ज BPM/SPM/PM यांनी नियमित Supervision ठेऊन पुढील दिवसांमध्ये "
           "कमीत कमी आर्टिकल deposit राहतील याची दक्षता घ्यावी.")
TEXT_B2 = ("या यादीतील ऑफिसेस मध्ये चालू महिन्याच्या ०१ तारखेपासून कालपर्यंत DSS% हा ९०% पेक्षा कमी "
           "आहे. ऑफिस चे इन्चार्ज BPM/SPM/PM यांनी नियमित Supervision ठेऊन पुढील दिवसांमध्ये 100% "
           "आर्टिकल DSS मधून वाटप होतील याची दक्षता घ्यावी.")
TEXT_B3 = ("या यादीतील ऑफिसेस मध्ये चालू महिन्याच्या ०१ तारखेपासून कालपर्यंत COD Digital Transaction "
           "% हे ९०% पेक्षा कमी आहे. ऑफिस चे इन्चार्ज BPM/SPM/PM यांनी नियमित Supervision ठेऊन "
           "पुढील दिवसांमध्ये जास्तीत जास्त COD पार्सल हे Digital Payment द्वारे पेड होतील याची "
           "दक्षता घ्यावी.")

PCT_FMT = '0.00%'
CNT_FMT = '#,##0'


def _fmts(wb):
    props = {}
    props['title'] = {'font_name': 'Arial', 'font_size': 16, 'bold': True, 'align': 'center', 'valign': 'vcenter'}
    props['record'] = {'font_name': 'Arial', 'font_size': 12, 'align': 'right', 'valign': 'vcenter'}
    props['marathi'] = {'font_name': 'Nirmala UI', 'font_size': 12, 'bold': True, 'text_wrap': True, 'valign': 'vcenter'}
    props['header'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'font_color': 'white',
                        'bg_color': NAVY, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID}
    props['body'] = {'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
                      'border': 1, 'border_color': GRID}
    props['body_left'] = {'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'valign': 'vcenter',
                           'border': 1, 'border_color': GRID}
    props['body_pct'] = {'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
                          'border': 1, 'border_color': GRID, 'num_format': PCT_FMT}
    props['body_cnt'] = {'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
                          'border': 1, 'border_color': GRID, 'num_format': CNT_FMT}
    props['total'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'bg_color': PEACH,
                       'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': GRID}
    props['total_pct'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'bg_color': PEACH,
                           'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID, 'num_format': PCT_FMT}
    props['total_cnt'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'bg_color': PEACH,
                           'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID, 'num_format': CNT_FMT}
    props['total_ctr'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'bg_color': PEACH,
                           'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID}
    props['section'] = {'font_name': 'Arial', 'font_size': 13, 'bold': True, 'font_color': 'white',
                         'bg_color': RED, 'align': 'center', 'valign': 'vcenter'}
    props['redbold'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'font_color': REDTEXT,
                         'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID}
    props['redbold_left'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'font_color': REDTEXT,
                              'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': GRID}
    props['redbold_pct'] = {'font_name': 'Arial', 'font_size': 12, 'bold': True, 'font_color': REDTEXT,
                             'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': GRID, 'num_format': PCT_FMT}
    props['note'] = {'font_name': 'Arial', 'font_size': 12, 'italic': True}
    props['plain'] = {'font_name': 'Arial', 'font_size': 12}

    # Threshold conditional-formatting fills (v9, new): standard Excel
    # "Good"/"Bad" cell-style colors, used for the red/green highlighting on
    # Divisional Summary, Excl BOs Total/Subtotal rows, and (text-only /
    # text+fill variants) the Defaulters sheets' COD Range % column.
    props['cf_green'] = {'bg_color': '#C6EFCE', 'font_color': '#006100'}
    props['cf_red'] = {'bg_color': '#FFC7CE', 'font_color': '#9C0006'}
    props['cf_red_text'] = {'font_color': '#9C0006'}

    # "_gb" (group border) variants: same as the base format but with a thick left
    # border, used on the first column of each metric group (Delivery/DSS/COD/LB)
    # in the grouped 2-row headers, so the groups are visually segregated.
    for base in ('header', 'body_pct', 'body_cnt', 'total_pct', 'total_cnt'):
        p = dict(props[base])
        p['left'] = 5
        p['left_color'] = GROUP_BORDER
        props[base + '_gb'] = p

    f = {name: wb.add_format(p) for name, p in props.items()}
    f['_props'] = props
    f['_wb'] = wb
    return f


def _variant(f, base_name, **overrides):
    """Build (and cache-free-create) a format that's the named base plus extra
    overrides merged in — used for one-off top/bottom border rows (group-header
    top edge, sub-header bottom edge, total-row top+bottom edge)."""
    p = dict(f['_props'][base_name])
    p.update(overrides)
    return f['_wb'].add_format(p)


def _color_scale(ws, first_row, last_row, col):
    if last_row < first_row:
        return
    ws.conditional_format(first_row, col, last_row, col, {
        'type': '3_color_scale',
        'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B',
        'min_type': 'min', 'mid_type': 'percentile', 'mid_value': 50, 'max_type': 'max',
    })


def _two_color_scale(ws, first_row, last_row, col):
    """Red -> near-white 2-color scale (no green/yellow) -- used on the
    Defaulters sheets' own Range % columns (Delivery, DSS), matching the
    manually-set reference formatting."""
    if last_row < first_row:
        return
    ws.conditional_format(first_row, col, last_row, col, {
        'type': '2_color_scale',
        'min_color': '#F8696B', 'max_color': '#FCFCFF',
    })


def _data_bar(ws, first_row, last_row, col, color='#FF555A'):
    if last_row < first_row:
        return
    ws.conditional_format(first_row, col, last_row, col, {
        'type': 'data_bar', 'bar_color': color,
    })


def _threshold_scale(ws, first_row, last_row, col_first, col_last, green_above, red_below,
                      green_fmt, red_fmt):
    """Threshold-based red/green highlighting (replaces the smooth 3-color
    gradient on Divisional Summary and on the Total/Subtotal rows of the Excl
    BOs sheets): green fill if strictly greater than green_above, red fill if
    strictly less than red_below -- values between the two thresholds get no
    highlight, same as the manually-set reference formatting."""
    if last_row < first_row:
        return
    ws.conditional_format(first_row, col_first, last_row, col_last, {
        'type': 'cell', 'criteria': '>', 'value': green_above, 'format': green_fmt,
    })
    ws.conditional_format(first_row, col_first, last_row, col_last, {
        'type': 'cell', 'criteria': '<', 'value': red_below, 'format': red_fmt,
    })


# Metric-specific thresholds for the red/green highlighting above (v9, new) --
# used on Divisional Summary and on the Excl BOs Total/Subtotal rows.
METRIC_THRESHOLDS = {
    'delivery': (0.95, 0.90),  # (green if >, red if <)
    'dss': (0.99, 0.95),
    'cod': (0.90, 0.80),
    'lb': (0.99, 0.95),
}


def _cod_range_defaulter_highlight(ws, first_row, last_row, col, red_fmt_text, red_fmt_fill):
    """Exact reference formatting for the Defaulters sheets' COD Digital Txn
    (Range) column: red TEXT (no fill) if < 60%, red text + red FILL if
    exactly 0% -- two distinct severity levels within an already-filtered
    (<90%) defaulters list."""
    if last_row < first_row:
        return
    ws.conditional_format(first_row, col, last_row, col, {
        'type': 'cell', 'criteria': '<', 'value': 0.6, 'format': red_fmt_text,
    })
    ws.conditional_format(first_row, col, last_row, col, {
        'type': 'cell', 'criteria': '=', 'value': 0, 'format': red_fmt_fill,
    })


def _autofit(ws, ncols, headers, data_rows_texts, min_w=10, max_w=45):
    for j in range(ncols):
        width = max([len(str(headers[j]))] + [len(str(t)) for t in data_rows_texts[j]], default=min_w)
        ws.set_column(j, j, max(min_w, min(max_w, width + 3)))


def _safe_sheet_name(name, used):
    """Excel sheet names: max 31 chars, none of  [ ] : * ? / \\ , and unique."""
    base = re.sub(r'[\[\]:*?/\\]', '-', name).strip()[:31] or 'Sheet'
    cand, i = base, 2
    while cand.lower() in used:
        suffix = f' ({i})'
        cand = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(cand.lower())
    return cand


def build_workbook(master_raw, df, excl_r, excl_sd, avail, single_date, range_start, range_end,
                    dp_r1_df, dp_r2_df, dp_sd_df, dss_r_df, dss_sd_df, lb_r_df, lb_sd_df,
                    cod_r_df, cod_sd_df, date_meta=None, division_name='Division',
                    record_prefix=None, findings=None, signature=None, idc_inputs=None):
    """
    master_raw: 5-col Master dataframe (office_id, office_name, sub_office, sub_division, office_type)
    df: full computed dataframe from engine.build_full (one row per Master office)
    avail: dict of 8 booleans from engine.build_full
    single_date, range_start, range_end: datetime.date objects (canonical, for title/record no.)
    date_meta: optional dict of the 9 individual file date selections, written to Validation Log
    *_df: the raw uploaded dataframes (or None), for the hidden Src_* mirror sheets
    division_name: from the Master ('Karad Division') — used in total-row labels
    record_prefix: office-record-number prefix (default derived from division_name, e.g. 'KRDDN')
    findings: engine validation findings, written to the hidden Validation Log
    idc_inputs: optional dict(dpr, dss, office_label) of *prepared* IDC frames (see idc.prepare). When given,
               an 'IDC Report' sheet (+ hidden Src_IDC_* sheets) is added to this same workbook.
    signature: list of text lines for the signature block under every report sheet
               (default: derived from division_name; blank lines are skipped)
    Returns: BytesIO
    """
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True})
    f = _fmts(wb)

    sub_divs = get_sub_divisions(df)
    sig_lines = clean_lines(signature if signature is not None else default_signature(division_name))
    f_sig = wb.add_format({'font_name': 'Arial', 'font_size': 12, 'bold': True, 'align': 'left', 'valign': 'vcenter'})

    def write_signature(sheet, last_row, start_col):
        """Signature block 2 blank rows below the table (text overflows to the right, like the IDC report)."""
        for k, t in enumerate(sig_lines):
            sheet.write(last_row + 3 + k, start_col, t, f_sig)
    division_name = (division_name or 'Division').strip()
    total_label = f'{division_name} Total'
    record_prefix = (record_prefix or default_record_prefix(division_name)).strip()
    record_date = (single_date + timedelta(days=1)).strftime('%d.%m.%y')
    record_no = f"No. {record_prefix}/G/Mails/MMU/Performance/{record_date}"
    single_str = single_date.strftime('%d.%m.%Y')
    title_locked = f"MMU Report Only S.O {single_str}"
    used_sheet_names = set()

    # Date-label strings for column headers (v9: real dates, never generic "(Range)"/"(Single Date)")
    # Raw Data uses the verbose "01 July to 28 July" form; every other sheet uses the compact
    # "01-28 July" form — matching the two styles already present in the manual-edit reference file.
    range_verbose = f"{range_start.strftime('%d %B')} to {range_end.strftime('%d %B')}"
    range_compact = f"{range_start.strftime('%d')}-{range_end.strftime('%d %B')}"
    single_label = single_date.strftime('%d %B')

    ds = build_divisional_summary(df, total_label=total_label)
    excl_tbl = build_excl_bo_table(df)
    defs_ = build_defaulters(df)
    exc = build_exceptional_defaulters(df)

    # ============================================================ hidden sheets ===
    ws_master = wb.add_worksheet('Master')
    ws_master.hide()
    headers = ['office_id', 'office_name', 'sub_office', 'sub_division', 'office_type']
    for j, h in enumerate(headers):
        ws_master.write(0, j, h, f['plain'])
    for i, row in enumerate(master_raw.itertuples(index=False), 1):
        ws_master.write(i, 0, row.office_id)
        ws_master.write(i, 1, row.office_name)
        ws_master.write(i, 2, row.sub_office)
        ws_master.write(i, 3, row.sub_division)
        ws_master.write(i, 4, row.office_type)

    def dump_hidden(name, src_df, cols):
        ws = wb.add_worksheet(name)
        ws.hide()
        if src_df is None:
            ws.write(0, 0, '(no file uploaded this run)', f['note'])
            return
        for j, h in enumerate(cols):
            ws.write(0, j, h, f['plain'])
        # reindex keeps every column in its fixed position (Raw Data's SUMIFS formulas
        # refer to columns by letter) even if an optional column is absent from the export.
        for i, row in enumerate(src_df.reindex(columns=cols).itertuples(index=False), 1):
            for j, v in enumerate(row):
                if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
                    ws.write_blank(i, j, None)
                else:
                    ws.write(i, j, v)

    dp_r_cols = ['office-id', 'office-name', 'postman-count', 'invoice-count', 'delivery-count',
                 'deposit-count', 'return-count', 'redirection-count', 'beat-diversion-count']
    dss_cols = ['office_id', 'office_name', 'event_date', 'dss_art_count', 'pdm_art_count',
                'Mobile Usage %', 'total_pdm_art_count', 'total_dss_art_count', 'total_dss_count']
    lb_r_cols = ['office-id', 'office-name', 'lb-location', 'total-letterboxes', 'total-letters',
                 'avg-cleared', 'total-cleared', 'working-days', 'clearance-percentage']
    lb_sd_cols = ['office-id', 'office-name', 'total-letterboxes', 'total-letters',
                  'avg-cleared', 'total-cleared', 'clearance-percentage']
    cod_cols = ['Circle Name', 'Region Name', 'Division Name', 'Sub Division Name', 'Office Name',
                'Office ID', 'Customer Name', 'Product Name', 'COD Cash Count', 'COD Cash Amount',
                'COD Digital Count', 'COD Digital Amount', 'Total COD Count', 'Total COD Amount',
                'Booking COD Amount', 'COD Collection Performance (%)']

    dump_hidden('Src_DP_R1', dp_r1_df, dp_r_cols)
    dump_hidden('Src_DP_R2', dp_r2_df, dp_r_cols)
    dump_hidden('Src_DP_SD', dp_sd_df, dp_r_cols)
    dump_hidden('Src_DSS_R', dss_r_df, dss_cols)
    dump_hidden('Src_DSS_SD', dss_sd_df, dss_cols)
    dump_hidden('Src_LB_R', lb_r_df, lb_r_cols)
    dump_hidden('Src_LB_SD', lb_sd_df, lb_sd_cols)
    dump_hidden('Src_CODD_R', cod_r_df, cod_cols)
    dump_hidden('Src_CODD_SD', cod_sd_df, cod_cols)

    ws_vlog = wb.add_worksheet('Validation Log')
    ws_vlog.hide()
    vheaders = ['Check', 'Detail', 'Result']
    for j, h in enumerate(vheaders):
        ws_vlog.write(0, j, h, f['plain'])
    def _dss_excluded(ex):
        names = [f"{r.office_name} ({r.office_id})" for r in ex.drop_duplicates('office_id').itertuples()
                 if r.office_id != 0]
        return ', '.join(names) if names else '(none)'

    n_dup_master = int(master_raw.office_id.duplicated().sum())
    vrows = [
        ('Master reconciliation',
         f"{len(master_raw)} offices; BPO={sum(master_raw.office_type=='BPO')}, "
         f"SPO={sum(master_raw.office_type=='SPO')}, HPO={sum(master_raw.office_type=='HPO')}; "
         f"sub-divisions={len(sub_divs)}; duplicate Office IDs={n_dup_master}",
         'PASS' if n_dup_master == 0 else 'FAIL'),
        ('DSS Range — rows excluded (office_id=0 / not in Master)',
         _dss_excluded(excl_r) if avail['dss_r'] else '(file not provided)', 'EXCLUDED, LOGGED'),
        ('DSS Single Date — rows excluded (office_id=0 / not in Master)',
         _dss_excluded(excl_sd) if avail['dss_sd'] else '(file not provided)', 'EXCLUDED, LOGGED'),
        ('Data availability this run', ', '.join(f"{k}={'YES' if v else 'MISSING'}" for k, v in avail.items()), 'INFO'),
        ('Canonical report dates (title/record no.)',
         f"Single Date={single_date}, Range Start={range_start}, Range End={range_end}", 'INFO'),
        ('Division / record-number prefix', f"{division_name}; No. {record_prefix}/G/Mails/MMU/Performance/…", 'INFO'),
    ]
    _lvl = {'pass': 'PASS', 'info': 'INFO', 'warning': 'WARNING', 'error': 'FAIL'}
    for x in (findings or []):
        vrows.append((x['check'], x['detail'], _lvl.get(x['level'], x['level'].upper())))
    if date_meta:
        vrows.append(('Per-file calendar-picker dates (never inferred from filenames)',
                       '; '.join(f"{k}={v}" for k, v in date_meta.items() if v is not None), 'INFO'))
    for i, (a, b, c) in enumerate(vrows, 1):
        ws_vlog.write(i, 0, a, f['plain'])
        ws_vlog.write(i, 1, b, f['plain'])
        ws_vlog.write(i, 2, c, f['plain'])
    ws_vlog.set_column(0, 0, 35)
    ws_vlog.set_column(1, 1, 90)
    ws_vlog.set_column(2, 2, 18)

    # ============================================================ RAW DATA ===
    ws = wb.add_worksheet('Raw Data')
    ws.hide_gridlines(2)
    ws.set_tab_color(NAVY)

    RD_HEADERS = ['Sr. No.', 'Office ID', 'Office Name', 'Sub Office Name', 'Sub Division Name', 'Office Type',
                  f'Invoice Count ({range_verbose})', f'Delivery % ({range_verbose})',
                  f'Invoice-Count ({single_label})', f'Delivery % ({single_label})',
                  f'DSS % ({range_verbose})', f'DSS % ({single_label})',
                  f'Total COD Count ({range_verbose})', f'COD % ({range_verbose})',
                  f'Total COD Count ({single_label})', f'COD % ({single_label})',
                  'Total Letter Boxes', f'LB % ({range_verbose})', f'LB % ({single_label})']
    NCOLS_RD = len(RD_HEADERS)  # 19 (A-S)
    ws.merge_range(0, 0, 0, NCOLS_RD - 4, title_locked, f['title'])
    ws.merge_range(0, NCOLS_RD - 3, 0, NCOLS_RD - 1, record_no, f['record'])
    ws.set_row(0, 24)
    ws.set_row(1, 24)  # blank spacer row
    header_wrap = _variant(f, 'header', text_wrap=True)
    for j, h in enumerate(RD_HEADERS):
        ws.write(2, j, h, header_wrap)
    ws.set_row(2, 30)

    master_sorted = master_raw.sort_values('office_id').reset_index(drop=True)
    df_by_id = df.set_index('office_id', drop=False)  # office_id is unique (enforced by engine)
    DATA_START = 3  # 0-indexed row 3 == Excel row 4 (title=1, blank=2, headers=3, data=4+)
    for i, row in enumerate(master_sorted.itertuples(index=False)):
        rr = DATA_START + i
        oid = row.office_id
        drow = df_by_id.loc[oid]
        ws.write(rr, 0, i + 1, f['body'])
        ws.write(rr, 1, oid, f['body'])
        ws.write(rr, 2, row.office_name, f['body'])
        ws.write(rr, 3, row.sub_office, f['body'])
        ws.write(rr, 4, row.sub_division, f['body'])
        ws.write(rr, 5, row.office_type, f['body'])

        excel_row = rr + 1  # 1-indexed for formula text
        b = f"$B{excel_row}"

        inv_r = (f"(SUMIFS(Src_DP_R1!$D:$D,Src_DP_R1!$A:$A,{b})"
                 f"+SUMIFS(Src_DP_R2!$D:$D,Src_DP_R2!$A:$A,{b}))")
        dep_r = (f"(SUMIFS(Src_DP_R1!$F:$F,Src_DP_R1!$A:$A,{b})"
                 f"+SUMIFS(Src_DP_R2!$F:$F,Src_DP_R2!$A:$A,{b}))")
        if avail['delivery_r']:
            ws.write_formula(rr, 6, f"={inv_r}", f['body_cnt'],
                              0 if pd.isna(drow['invoice_r']) else drow['invoice_r'])  # SUMIFS with no match = 0 in Excel
            formula = f'=IF({inv_r}=0,"",({inv_r}-{dep_r})/{inv_r})'
            val = drow['delivery_pct_r']
            ws.write_formula(rr, 7, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 6, None, f['body'])
            ws.write_blank(rr, 7, None, f['body'])

        inv_sd = f"SUMIFS(Src_DP_SD!$D:$D,Src_DP_SD!$A:$A,{b})"
        dep_sd = f"SUMIFS(Src_DP_SD!$F:$F,Src_DP_SD!$A:$A,{b})"
        if avail['delivery_sd']:
            ws.write_formula(rr, 8, f"={inv_sd}", f['body_cnt'],
                              0 if pd.isna(drow['invoice_sd']) else drow['invoice_sd'])  # SUMIFS with no match = 0 in Excel
            formula = f'=IF({inv_sd}=0,"",({inv_sd}-{dep_sd})/{inv_sd})'
            val = drow['delivery_pct_sd']
            ws.write_formula(rr, 9, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 8, None, f['body'])
            ws.write_blank(rr, 9, None, f['body'])

        if avail['dss_r']:
            pdm_r = f"SUMIFS(Src_DSS_R!$G:$G,Src_DSS_R!$A:$A,{b})"
            dss_r = f"SUMIFS(Src_DSS_R!$H:$H,Src_DSS_R!$A:$A,{b})"
            formula = f'=IF({pdm_r}=0,"",{dss_r}/{pdm_r})'
            val = drow['dss_pct_r']
            ws.write_formula(rr, 10, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 10, None, f['body'])

        if avail['dss_sd']:
            pdm_sd = f"SUMIFS(Src_DSS_SD!$G:$G,Src_DSS_SD!$A:$A,{b})"
            dss_sd = f"SUMIFS(Src_DSS_SD!$H:$H,Src_DSS_SD!$A:$A,{b})"
            formula = f'=IF({pdm_sd}=0,"",{dss_sd}/{pdm_sd})'
            val = drow['dss_pct_sd']
            ws.write_formula(rr, 11, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 11, None, f['body'])

        tot_r = f"SUMIFS(Src_CODD_R!$M:$M,Src_CODD_R!$F:$F,{b})"
        dig_r = f"SUMIFS(Src_CODD_R!$K:$K,Src_CODD_R!$F:$F,{b})"
        if avail['cod_r']:
            ws.write_formula(rr, 12, f"={tot_r}", f['body_cnt'],
                              0 if pd.isna(drow['total_r']) else drow['total_r'])  # SUMIFS with no match = 0 in Excel
            formula = f'=IF({tot_r}=0,"",{dig_r}/{tot_r})'
            val = drow['cod_pct_r']
            ws.write_formula(rr, 13, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 12, None, f['body'])
            ws.write_blank(rr, 13, None, f['body'])

        tot_sd = f"SUMIFS(Src_CODD_SD!$M:$M,Src_CODD_SD!$F:$F,{b})"
        dig_sd = f"SUMIFS(Src_CODD_SD!$K:$K,Src_CODD_SD!$F:$F,{b})"
        if avail['cod_sd']:
            ws.write_formula(rr, 14, f"={tot_sd}", f['body_cnt'],
                              0 if pd.isna(drow['total_sd']) else drow['total_sd'])  # SUMIFS with no match = 0 in Excel
            formula = f'=IF({tot_sd}=0,"",{dig_sd}/{tot_sd})'
            val = drow['cod_pct_sd']
            ws.write_formula(rr, 15, formula, f['body_pct'], '' if pd.isna(val) else val)
        else:
            ws.write_blank(rr, 14, None, f['body'])
            ws.write_blank(rr, 15, None, f['body'])

        is_bpo = row.office_type == 'BPO'
        # Total Letter Boxes
        if is_bpo or not (avail['lb_r'] or avail['lb_sd']):
            ws.write_blank(rr, 16, None, f['body'])
        else:
            val = drow['total_letterboxes']
            if avail['lb_sd']:
                cnt_sd = f"COUNTIFS(Src_LB_SD!$A:$A,{b})"
                if avail['lb_r']:
                    cnt_r = f"COUNTIFS(Src_LB_R!$A:$A,{b})"
                    formula = (f'=IF({cnt_sd}>0,SUMIFS(Src_LB_SD!$C:$C,Src_LB_SD!$A:$A,{b}),'
                               f'IF({cnt_r}>0,AVERAGEIFS(Src_LB_R!$D:$D,Src_LB_R!$A:$A,{b}),""))')
                else:
                    formula = f'=IF({cnt_sd}>0,SUMIFS(Src_LB_SD!$C:$C,Src_LB_SD!$A:$A,{b}),"")'
            else:
                cnt_r = f"COUNTIFS(Src_LB_R!$A:$A,{b})"
                formula = f'=IF({cnt_r}>0,AVERAGEIFS(Src_LB_R!$D:$D,Src_LB_R!$A:$A,{b}),"")'
            ws.write_formula(rr, 16, formula, f['body_cnt'], '' if pd.isna(val) else val)

        if is_bpo or not avail['lb_r']:
            ws.write_blank(rr, 17, None, f['body'])
        else:
            box_r = f"SUMIFS(Src_LB_R!$D:$D,Src_LB_R!$A:$A,{b})"
            clr_r = f"SUMIFS(Src_LB_R!$G:$G,Src_LB_R!$A:$A,{b})"
            formula = f'=IF({box_r}=0,"",{clr_r}/{box_r})'
            val = drow['lb_pct_r']
            ws.write_formula(rr, 17, formula, f['body_pct'], '' if pd.isna(val) else val)

        if is_bpo or not avail['lb_sd']:
            ws.write_blank(rr, 18, None, f['body'])
        else:
            box_sd = f"SUMIFS(Src_LB_SD!$C:$C,Src_LB_SD!$A:$A,{b})"
            clr_sd = f"SUMIFS(Src_LB_SD!$F:$F,Src_LB_SD!$A:$A,{b})"
            formula = f'=IF({box_sd}=0,"",{clr_sd}/{box_sd})'
            val = drow['lb_pct_sd']
            ws.write_formula(rr, 18, formula, f['body_pct'], '' if pd.isna(val) else val)

    DATA_END = DATA_START + len(master_sorted) - 1
    for j in [7, 9, 10, 11, 13, 15, 17, 18]:
        _color_scale(ws, DATA_START, DATA_END, j)
    ws.freeze_panes(3, 6)  # lock rows 1-3 (title, blank, headers) and columns A-F
    widths = [8.57, 11.57, 29, 22, 22.43, 13.71, 18, 22.57, 22.57, 28,
              18.43, 23.86, 23.86, 18.71, 18.71, 24.14, 21.43, 16.57, 22.14]
    for j, w in enumerate(widths):
        ws.set_column(j, j, w)

    # ==================================================== EXCL BOs COMBINED ===
    EXCL_HEADERS = ['Sr. No.', 'Office Name', 'Delivery % (R)', 'Delivery % (SD)', 'DSS % (R)',
                     'DSS % (SD)', 'COD % (R)', 'COD % (SD)', 'Total Letter Boxes', 'LB % (R)', 'LB % (SD)']
    NCOLS_EXCL = len(EXCL_HEADERS)

    def write_excl_row(ws, rr, sr_no, name, r, total=False):
        vals = [sr_no, name, r['delivery_pct_r'], r['delivery_pct_sd'], r['dss_pct_r'], r['dss_pct_sd'],
                r['cod_pct_r'], r['cod_pct_sd'], r['total_letterboxes'], r['lb_pct_r'], r['lb_pct_sd']]
        group_start_cols = {2, 4, 6, 8}  # first column of Delivery/DSS/COD/LB groups
        box = {'top': 5, 'top_color': GROUP_BORDER, 'bottom': 5, 'bottom_color': GROUP_BORDER} if total else {}
        for j, v in enumerate(vals):
            gb = j in group_start_cols
            if j == 0:
                base = 'total_ctr' if total else 'body'
            elif j == 1:
                base = 'total' if total else 'body_left'
            elif j == 8:
                base = 'total_cnt_gb' if total else ('body_cnt_gb' if gb else 'body_cnt')
            else:
                if gb:
                    base = 'total_pct_gb' if total else 'body_pct_gb'
                else:
                    base = 'total_pct' if total else 'body_pct'
            fmt = _variant(f, base, **box) if total else f[base]
            if v is None or (isinstance(v, float) and pd.isna(v)):
                ws.write_blank(rr, j, None, fmt)
            else:
                ws.write(rr, j, v, fmt)

    def write_grouped_header_excl(ws, r):
        """2-row grouped header: metric name row + date sub-label row. Returns next row.
        Thick left border marks the start of each metric group; thick top border
        caps the group-name row and thick bottom border caps the date-label row,
        matching the header/data segregation seen in the manual-edit reference file."""
        group_start_cols = {2, 4, 6, 8}
        TOP = {'top': 5, 'top_color': GROUP_BORDER}
        BOTTOM = {'bottom': 5, 'bottom_color': GROUP_BORDER, 'text_wrap': True}
        TOP_BOTTOM = {**TOP, **BOTTOM}

        def hfmt(col, extra):
            base = 'header_gb' if col in group_start_cols else 'header'
            return _variant(f, base, **extra)

        # Sr. No. / Office Name span both header rows, so they need top AND bottom.
        ws.merge_range(r, 0, r + 1, 0, 'Sr. No.', hfmt(0, TOP_BOTTOM))
        ws.merge_range(r, 1, r + 1, 1, 'Office Name', hfmt(1, TOP_BOTTOM))
        # Group-name row: top border only (bottom edge is internal, shared with the row below).
        ws.merge_range(r, 2, r, 3, 'Delivery % (R)', hfmt(2, TOP))
        ws.merge_range(r, 4, r, 5, 'DSS Usage %', hfmt(4, TOP))
        ws.merge_range(r, 6, r, 7, 'COD Digital Transaction %', hfmt(6, TOP))
        ws.merge_range(r, 8, r, 10, 'Letter Box Clearance %', hfmt(8, TOP))
        r += 1
        # Date sub-label row: bottom border + wrap (this is the header/data boundary).
        ws.write(r, 2, range_compact, hfmt(2, BOTTOM))
        ws.write(r, 3, single_label, hfmt(3, BOTTOM))
        ws.write(r, 4, range_compact, hfmt(4, BOTTOM))
        ws.write(r, 5, single_label, hfmt(5, BOTTOM))
        ws.write(r, 6, range_compact, hfmt(6, BOTTOM))
        ws.write(r, 7, single_label, hfmt(7, BOTTOM))
        ws.write(r, 8, 'Total Letter Boxes', hfmt(8, BOTTOM))
        ws.write(r, 9, range_compact, hfmt(9, BOTTOM))
        ws.write(r, 10, single_label, hfmt(10, BOTTOM))
        return r + 1

    ws2 = wb.add_worksheet('Excl BOs - Combined')
    ws2.hide_gridlines(2)
    ws2.set_tab_color(RED)
    ws2.merge_range(0, 0, 0, NCOLS_EXCL - 1, title_locked, f['title'])
    ws2.merge_range(1, 0, 1, NCOLS_EXCL - 1, record_no, f['record'])
    ws2.merge_range(2, 0, 2, NCOLS_EXCL - 1, TEXT_A, _variant(f, 'marathi', align='center'))
    ws2.set_row(2, 45)
    data_start = write_grouped_header_excl(ws2, 3)
    for i, row in enumerate(excl_tbl.itertuples(), 0):
        write_excl_row(ws2, data_start + i, i + 1, row.office_name, row._asdict())
    data_end = data_start + len(excl_tbl) - 1
    total_row_idx = data_end + 1
    totals = weighted_row(excl_tbl)
    totals['total_letterboxes'] = excl_tbl['total_letterboxes'].sum(skipna=True)
    write_excl_row(ws2, total_row_idx, '', f'{total_label} (SPO+HPO)', totals, total=True)
    for j in [2, 3, 4, 5, 6, 7, 9, 10]:
        _color_scale(ws2, data_start, data_end, j)
    # Total row (v9, new): threshold red/green highlighting per metric group,
    # replacing the smooth gradient (which is data-row-only, unchanged above).
    for (c1, c2), key in [((2, 3), 'delivery'), ((4, 5), 'dss'), ((6, 7), 'cod'), ((9, 10), 'lb')]:
        green_above, red_below = METRIC_THRESHOLDS[key]
        _threshold_scale(ws2, total_row_idx, total_row_idx, c1, c2, green_above, red_below,
                          f['cf_green'], f['cf_red'])
    write_signature(ws2, total_row_idx, 4)
    # freeze_panes intentionally removed (unfrozen) per request
    for j, w in enumerate([7, 26, 13, 13, 13, 13, 13, 13, 15, 13, 13]):
        ws2.set_column(j, j, w)

    # ================================================= EXCL BOs SUB DIVISION ===
    ws3 = wb.add_worksheet('Excl BOs - Sub Division')
    ws3.hide_gridlines(2)
    ws3.set_tab_color(RED)

    r = 0
    first_block_data_start = None
    for sd in sub_divs:
        ws3.merge_range(r, 0, r, NCOLS_EXCL - 1, title_locked, f['title'])
        ws3.set_row(r, 24)
        r += 1
        ws3.merge_range(r, 0, r, NCOLS_EXCL - 1, record_no, f['record'])
        r += 1
        ws3.merge_range(r, 0, r, NCOLS_EXCL - 1, TEXT_A, _variant(f, 'marathi', align='center'))
        ws3.set_row(r, 45)
        r += 1
        ws3.merge_range(r, 0, r, NCOLS_EXCL - 1, sd, f['section'])
        r += 1
        r = write_grouped_header_excl(ws3, r)
        if first_block_data_start is None:
            first_block_data_start = r
        sub = excl_tbl[excl_tbl['sub_division'] == sd]
        block_start = r
        for i, row in enumerate(sub.itertuples(), 0):
            write_excl_row(ws3, r, i + 1, row.office_name, row._asdict())
            r += 1
        block_end = r - 1
        subtotal = weighted_row(sub)
        subtotal['total_letterboxes'] = sub['total_letterboxes'].sum(skipna=True)
        write_excl_row(ws3, r, '', f'{sd} Subtotal', subtotal, total=True)
        for j in [2, 3, 4, 5, 6, 7, 9, 10]:
            _color_scale(ws3, block_start, block_end, j)
        # Subtotal row (v9, new): same threshold red/green highlighting as the
        # Combined sheet's Total row, applied independently per block.
        for (c1, c2), key in [((2, 3), 'delivery'), ((4, 5), 'dss'), ((6, 7), 'cod'), ((9, 10), 'lb')]:
            green_above, red_below = METRIC_THRESHOLDS[key]
            _threshold_scale(ws3, r, r, c1, c2, green_above, red_below, f['cf_green'], f['cf_red'])
        r += 2
    write_signature(ws3, r - 2, 4)
    # freeze_panes intentionally removed (unfrozen) per request
    for j, w in enumerate([7, 26, 13, 13, 13, 13, 13, 13, 15, 13, 13]):
        ws3.set_column(j, j, w)

    # ========================================================= DEFAULTERS ===
    marathi_red = _variant(f, 'marathi', font_color=REDTEXT)
    header_wrap = _variant(f, 'header', text_wrap=True)

    def write_marathi_highlighted(ws, row, ncols, text, highlights, base_fmt, red_fmt):
        """Write a Marathi paragraph in a merged cell, with specific substrings
        (e.g. 'डिलीवरी%', 'DSS%', 'COD Digital Transaction %') colored red for
        easy visual distinction. Falls back to a plain write if none are found."""
        segments = []
        pos = 0
        while pos < len(text):
            best_idx, best_hl = None, None
            for hl in highlights:
                idx = text.find(hl, pos)
                if idx != -1 and (best_idx is None or idx < best_idx):
                    best_idx, best_hl = idx, hl
            if best_idx is None:
                segments.append((base_fmt, text[pos:]))
                break
            if best_idx > pos:
                segments.append((base_fmt, text[pos:best_idx]))
            segments.append((red_fmt, best_hl))
            pos = best_idx + len(best_hl)

        if len(segments) <= 1:
            ws.merge_range(row, 0, row, ncols - 1, text, base_fmt)
            return
        ws.merge_range(row, 0, row, ncols - 1, '', base_fmt)
        parts = []
        for fmt, txt in segments:
            parts.append(fmt)
            parts.append(txt)
        ws.write_rich_string(row, 0, *parts, base_fmt)

    def write_def_table_v2(ws, r, label, marathi, highlight_words, columns_spec, sub_df, cat_missing, cf_map=None):
        ncols = len(columns_spec)
        ws.merge_range(r, 0, r, ncols - 1, label, f['section'])
        r += 1
        write_marathi_highlighted(ws, r, ncols, marathi, highlight_words, f['marathi'], marathi_red)
        ws.set_row(r, 45)
        r += 1
        for j, spec in enumerate(columns_spec):
            ws.write(r, j, spec['header'], header_wrap)
        r += 1
        data_start = r
        if cat_missing:
            ws.merge_range(r, 0, r, ncols - 1,
                            'Range % data not provided this run — defaulter list unavailable.', f['note'])
            return r + 3, data_start
        if len(sub_df) == 0:
            ws.merge_range(r, 0, r, ncols - 1, 'No defaulter offices for this parameter.', f['note'])
            return r + 3, data_start
        for i, row in enumerate(sub_df.itertuples(), 0):
            rr = r + i
            for j, spec in enumerate(columns_spec):
                kind = spec['kind']
                if kind == 'idx':
                    v = i + 1
                    fmt = f['body']
                elif kind == 'id':
                    v = row.office_id
                    fmt = f['body']
                elif kind == 'name':
                    v = row.office_name
                    fmt = f['body_left']
                elif kind == 'type':
                    v = row.office_type
                    fmt = f['body']
                elif kind == 'pct':
                    v = getattr(row, spec['field'])
                    fmt = f['body_pct']
                else:  # cnt
                    v = getattr(row, spec['field'])
                    fmt = f['body_cnt']
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    ws.write_blank(rr, j, None, fmt)
                else:
                    ws.write(rr, j, v, fmt)
        data_end = r + len(sub_df) - 1
        cf_map = cf_map or {}
        for j, spec in enumerate(columns_spec):
            if j in cf_map:
                cf_map[j](ws, data_start, data_end, j)
            elif spec['kind'] == 'pct':
                _color_scale(ws, data_start, data_end, j)
        return data_end + 2, data_start

    IDX = {'header': 'Sr. No.', 'kind': 'idx'}
    ID = {'header': 'Office ID', 'kind': 'id'}
    NAME = {'header': 'Office Name', 'kind': 'name'}
    TYPE = {'header': 'Office Type', 'kind': 'type'}

    # One Defaulters sheet per sub-division found in the Master (name trimmed to Excel's 31-char limit).
    name_map = {sd: _safe_sheet_name(f'Defaulters - {sd}', used_sheet_names) for sd in sub_divs}
    MAX_DEF_COLS = 8  # COD table is the widest (Sr,ID,Name,Type + 4 value cols)
    for sd in sub_divs:
        wsd = wb.add_worksheet(name_map[sd])
        wsd.hide_gridlines(2)
        wsd.set_tab_color(ORANGE)
        title = f"MMU Report Defaulters (<90%) - {sd} - {single_str}"
        wsd.merge_range(0, 0, 0, MAX_DEF_COLS - 1, title, f['title'])
        # office record number row intentionally removed per request — not needed on these sheets
        r = 1

        delivery_cols = [IDX, ID, NAME, TYPE,
                          {'header': f'Delivery % ({range_compact})', 'kind': 'pct', 'field': 'delivery_pct_r'},
                          {'header': f'Articles in Deposit ({single_label})', 'kind': 'cnt', 'field': 'deposit_sd'},
                          {'header': f'Delivery % ({single_label})', 'kind': 'pct', 'field': 'delivery_pct_sd'}]
        # v9 CF (matches the manually-set reference exactly): Range % gets a
        # 2-color (red->white) scale instead of the usual 3-color gradient;
        # the Single Date raw-count column (Articles in Deposit) gets a data
        # bar it didn't have before; Single Date % keeps the default gradient.
        delivery_cf = {4: lambda ws, s, e, c: _two_color_scale(ws, s, e, c),
                        5: lambda ws, s, e, c: _data_bar(ws, s, e, c)}
        r, ds1 = write_def_table_v2(wsd, r, 'Delivery % Defaulters (< 90%)', TEXT_B1, ['डिलीवरी%'],
                                     delivery_cols, defs_[sd]['delivery'], not avail['delivery_r'],
                                     cf_map=delivery_cf)

        dss_cols = [IDX, ID, NAME, TYPE,
                    {'header': f'DSS Usage ({range_compact})', 'kind': 'pct', 'field': 'dss_pct_r'},
                    {'header': f'Articles Invoiced on {single_label}', 'kind': 'cnt', 'field': 'pdm_sd'},
                    {'header': f'DSS Usage ({single_label})', 'kind': 'pct', 'field': 'dss_pct_sd'}]
        # v9 CF: Range % gets the same 2-color scale as Delivery; the count
        # column (Articles Invoiced) stays unformatted; Single Date % keeps
        # the default gradient.
        dss_cf = {4: lambda ws, s, e, c: _two_color_scale(ws, s, e, c)}
        r, _ = write_def_table_v2(wsd, r, 'DSS % Defaulters (< 90%)', TEXT_B2, ['DSS%'],
                                   dss_cols, defs_[sd]['dss'], not avail['dss_r'], cf_map=dss_cf)

        cod_cols = [IDX, ID, NAME, TYPE,
                    {'header': f'Total COD Articles Received ({range_compact})', 'kind': 'cnt', 'field': 'total_r'},
                    {'header': f'COD Digital Txn ({range_compact})', 'kind': 'pct', 'field': 'cod_pct_r'},
                    {'header': f'Total COD Articles Received ({single_label})', 'kind': 'cnt', 'field': 'total_sd'},
                    {'header': f'COD Digital Txn ({single_label})', 'kind': 'pct', 'field': 'cod_pct_sd'}]
        # v9 CF: Range raw-count column (Total COD Articles Received) gets a
        # data bar; Range % gets the two-tier red-text-only (<60%) / red-fill
        # (=0%) highlight instead of a gradient; Single Date % keeps the
        # default gradient.
        cod_cf = {4: lambda ws, s, e, c: _data_bar(ws, s, e, c),
                  5: lambda ws, s, e, c: _cod_range_defaulter_highlight(
                      ws, s, e, c, f['cf_red_text'], f['cf_red'])}
        r, _ = write_def_table_v2(wsd, r, 'COD Digital Transaction % Defaulters (< 90%)', TEXT_B3,
                                   ['COD Digital Transaction %'],
                                   cod_cols, defs_[sd]['cod'], not avail['cod_r'], cf_map=cod_cf)

        write_signature(wsd, r - 1, 4)
        # freeze_panes intentionally removed (unfrozen) per request
        for j, w in enumerate([7, 11, 26, 11, 20, 22, 20, 22]):
            wsd.set_column(j, j, w)

    # ===================================================== DIVISIONAL SUMMARY ===
    NCOLS_DS = 10
    ws7 = wb.add_worksheet('Divisional Summary')
    ws7.hide_gridlines(2)
    ws7.set_tab_color(PURPLE)
    # This sheet also covers the B.Os, so it must not be titled "Only S.O" (that title was being
    # corrected by hand after every run).
    ws7.merge_range(0, 0, 0, NCOLS_DS - 1, f"MMU Report {single_str} Divisional Summary", f['title'])
    ws7.merge_range(1, 0, 1, NCOLS_DS - 1, record_no, f['record'])
    ws7.merge_range(2, 0, 2, NCOLS_DS - 1, TEXT_A, f['marathi'])
    ws7.set_row(2, 45)
    r = 3

    def write_grouped_header_ds(ws, r):
        group_start_cols = {2, 4, 6, 8}
        TOP = {'top': 5, 'top_color': GROUP_BORDER}
        BOTTOM = {'bottom': 5, 'bottom_color': GROUP_BORDER}
        TOP_BOTTOM = {**TOP, **BOTTOM}

        def hfmt(col, extra):
            base = 'header_gb' if col in group_start_cols else 'header'
            return _variant(f, base, **extra)

        ws.merge_range(r, 0, r + 1, 0, 'Sr. No.', hfmt(0, TOP_BOTTOM))
        ws.merge_range(r, 1, r + 1, 1, 'Sub Division', hfmt(1, TOP_BOTTOM))
        ws.merge_range(r, 2, r, 3, 'Delivery %', hfmt(2, TOP))
        ws.merge_range(r, 4, r, 5, 'DSS Usage %', hfmt(4, TOP))
        ws.merge_range(r, 6, r, 7, 'COD Digital Txn %', hfmt(6, TOP))
        ws.merge_range(r, 8, r, 9, 'LB Clearance %', hfmt(8, TOP))
        r += 1
        for c in (2, 4, 6, 8):
            ws.write(r, c, range_compact, hfmt(c, BOTTOM))
            ws.write(r, c + 1, single_label, hfmt(c + 1, BOTTOM))
        return r + 1

    table_titles = {'excl_bo': 'Excluding B.Os (SPO + HPO only)', 'only_bo': 'Only B.Os (BPO only)',
                    'incl_bo': f'Including B.Os (All {len(df)} Offices)'}
    first_ds_data = None
    for key in ['excl_bo', 'only_bo', 'incl_bo']:
        ws7.merge_range(r, 0, r, NCOLS_DS - 1, table_titles[key], f['section'])
        r += 1
        r = write_grouped_header_ds(ws7, r)
        if first_ds_data is None:
            first_ds_data = r
        block_start = r
        rows = ds[key]
        group_start_cols = {2, 4, 6, 8}
        box = {'top': 5, 'top_color': GROUP_BORDER, 'bottom': 5, 'bottom_color': GROUP_BORDER, 'text_wrap': True}
        for i, row in enumerate(rows, 0):
            is_total = bool(row['is_total'])
            vals = [i + 1 if not is_total else '', row['sub_division'], row['delivery_pct_r'], row['delivery_pct_sd'],
                    row['dss_pct_r'], row['dss_pct_sd'], row['cod_pct_r'], row['cod_pct_sd'],
                    row['lb_pct_r'], row['lb_pct_sd']]
            for j, v in enumerate(vals):
                gb = j in group_start_cols
                if is_total:
                    if j == 0:
                        base = 'total_ctr'
                    elif j == 1:
                        base = 'total'
                    else:
                        base = 'total_pct_gb' if gb else 'total_pct'
                    fmt = _variant(f, base, **box)
                else:
                    if j == 0:
                        fmt = f['body']
                    elif j == 1:
                        fmt = f['body_left']
                    else:
                        fmt = f['body_pct_gb'] if gb else f['body_pct']
                if v is None or (isinstance(v, float) and pd.isna(v)) or v == '':
                    if v == '' and j == 0:
                        ws7.write_blank(r + i, j, None, fmt)
                    elif isinstance(v, str):
                        ws7.write(r + i, j, v, fmt)
                    else:
                        ws7.write_blank(r + i, j, None, fmt)
                else:
                    ws7.write(r + i, j, v, fmt)
        # Threshold red/green highlighting (v9, replaces the smooth gradient
        # entirely on this sheet) -- covers the sub-division rows AND the
        # Division Total row, per metric group. LB Clearance % is skipped for
        # the "Only B.Os" table since it's blank for every BPO row there.
        full_end = r + len(rows) - 1
        groups = [((2, 3), 'delivery'), ((4, 5), 'dss'), ((6, 7), 'cod')]
        if key != 'only_bo':
            groups.append(((8, 9), 'lb'))
        for (c1, c2), mkey in groups:
            green_above, red_below = METRIC_THRESHOLDS[mkey]
            _threshold_scale(ws7, block_start, full_end, c1, c2, green_above, red_below,
                              f['cf_green'], f['cf_red'])
        r += len(rows) + 1
    write_signature(ws7, r - 2, 4)
    # freeze_panes intentionally removed (unfrozen) per request
    for j, w in enumerate([7, 18, 13, 13, 13, 13, 13, 13, 13, 13]):
        ws7.set_column(j, j, w)

    # ============================================================ IDC REPORT (sits right after Divisional Summary) ===
    if idc_inputs:
        import idc as _idc
        _idc.write_idc_sheets(wb, idc_inputs['dpr'], idc_inputs['dss'], idc_inputs['office_label'], record_prefix,
                              single_date, range_start, range_end, None, sig_lines, standalone=False)

    # ================================================== EXCEPTIONAL DEFAULTERS ===
    ED_HEADERS = ['Sr. No.', 'Office ID', 'Office Name', 'Sub Division', 'Office Type', 'Parameter', 'Value %']
    NCOLS_ED = len(ED_HEADERS)
    ws8 = wb.add_worksheet('Exceptional Defaulters')
    ws8.hide_gridlines(2)
    ws8.set_tab_color(RED)
    title = f"MMU Report Exceptional Defaulters (< 30%) - {single_str}"
    ws8.merge_range(0, 0, 0, NCOLS_ED - 1, title, f['title'])
    ws8.merge_range(1, 0, 1, NCOLS_ED - 1, record_no, f['record'])
    header_row = 2
    for j, h in enumerate(ED_HEADERS):
        ws8.write(header_row, j, h, f['header'])
    data_start = header_row + 1
    for i, row in enumerate(exc.itertuples(), 0):
        rr = data_start + i
        parameter_label = row.parameter.replace('(Range)', f'({range_compact})').replace('(Single Date)', f'({single_label})')
        vals = [i + 1, row.office_id, row.office_name, row.sub_division, row.office_type, parameter_label, row.value]
        for j, v in enumerate(vals):
            fmt = f['redbold_left'] if j in (2, 5) else (f['redbold_pct'] if j == 6 else f['redbold'])
            ws8.write(rr, j, v, fmt)
    data_end = data_start + len(exc) - 1
    if len(exc) == 0:
        ws8.merge_range(data_start, 0, data_start, NCOLS_ED - 1,
                         'No offices below the 30% exceptional-defaulter threshold.', f['note'])
        data_end = data_start
    else:
        _color_scale(ws8, data_start, data_end, 6)
    write_signature(ws8, data_end, 3)
    ws8.freeze_panes(3, 0)  # A4
    for j, w in enumerate([7, 11, 26, 16, 11, 26, 10]):
        ws8.set_column(j, j, w)

    # ===================================================== sheet order/active ===
    ws.activate()

    wb.close()
    buf.seek(0)
    return buf
