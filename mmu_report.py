"""
MMU Hub Report — Streamlit webapp
Department of Posts — Mail Motor / Mobile Unit performance monitoring.

Division-agnostic: upload your Division's Office Master File (standard format) and
the header, sub-divisions, labels and report layout all follow from it.

- Master file: uploaded per session (optionally saved as this app's default)
- India Post branded header
- 9 named file placeholders, each with its own calendar picker(s) — dates are
  NEVER inferred from filenames, only from the pickers
- Every file is validated before anything is built (columns, duplicates,
  offices missing from the Master, swapped Range/Single-Date files, bounds)
- "Generate" -> if files missing, warn and ask "still generate the report?"
- Live Divisional Summary widget (Range + Single Date, all 4 KPIs)

Run locally:   streamlit run app.py
Deploy:        Streamlit Community Cloud (no system packages needed).
"""
import base64
import datetime
import html
import io
import mimetypes
import pandas as pd
import streamlit as st

import engine
import idc
from engine import (read_csv_any, read_master_full, build_full, build_divisional_summary,
                     save_master_file, load_persisted_master, save_logo_file, load_persisted_logo,
                     run_validations, check_percentage_bounds, has_errors, get_sub_divisions,
                     default_record_prefix)
from workbook import build_workbook

from ui import (parse_master, active_master_bytes, header_text, render_header, signature_inputs)

master_bytes, master_source = active_master_bytes()
master, master_info, master_findings = (parse_master(master_bytes) if master_bytes else (None, {}, []))
division_name = master_info.get('division_name', '') if master is not None else ''
header_title, header_sub = header_text(master, master_info)
render_header(header_title, header_sub)

# ------------------------------------------------------------------ state ---
if 'confirm_missing' not in st.session_state:
    st.session_state.confirm_missing = False
if 'generated_buf' not in st.session_state:
    st.session_state.generated_buf = None
    st.session_state.generated_sig = None


def show_findings(findings, expand_warnings=True):
    """Render validation findings: errors and warnings visibly, passes tucked away."""
    for x in findings:
        if x['level'] == 'error':
            st.error(f"**{x['check']}** — {x['detail']}")
    warns = [x for x in findings if x['level'] == 'warning']
    if warns:
        with st.expander(f"⚠️ {len(warns)} warning(s) — review before using the report", expanded=expand_warnings):
            for x in warns:
                st.warning(f"**{x['check']}** — {x['detail']}")
    ok = [x for x in findings if x['level'] in ('pass', 'info')]
    if ok:
        with st.expander(f"✅ {len([x for x in ok if x['level'] == 'pass'])} check(s) passed"):
            for x in ok:
                st.markdown(f"{'✅' if x['level'] == 'pass' else 'ℹ️'} **{x['check']}** — {x['detail']}")


# ----------------------------------------------------------- Master file ---
st.subheader("📋 Office Master File")
mcol1, mcol2 = st.columns([3, 2])
with mcol1:
    if master is not None:
        counts = master_info.get('type_counts', {})
        st.success(
            f"Using {master_source}. **{division_name or 'Division'}** — {master_info['n_offices']} offices "
            f"({', '.join(f'{v} {k}' for k, v in sorted(counts.items(), key=lambda kv: -kv[1]))}); "
            f"sub-divisions: {', '.join(master_info['sub_divisions'])}.")
        if master_source and 'default' in master_source:
            st.caption("This is the app's shared default. To use a different Division's Master, upload it on the "
                       "right — it applies to your session only and won't affect other users.")
    elif master_bytes:
        for x in master_findings:
            if x['level'] == 'error':
                st.error(f"**{x['check']}** — {x['detail']}")
    else:
        st.warning("No Master File yet — upload your Division's Office Master File to begin "
                   "(standard format: Office ID, Office Name, Sub Division Name, Office Type Code, …).")
with mcol2:
    with st.expander("Upload / Replace Master File", expanded=master is None):
        new_master = st.file_uploader("Office Master File (.xlsx)", type=["xlsx"], key="master_upload")
        if new_master is not None:
            data = new_master.getvalue()
            m_new, info_new, f_new = parse_master(data)
            if m_new is None:
                for x in f_new:
                    if x['level'] == 'error':
                        st.error(f"**{x['check']}** — {x['detail']}")
            else:
                st.success(f"Looks good: {info_new.get('division_name') or 'Division'} — "
                           f"{info_new['n_offices']} offices, {len(info_new['sub_divisions'])} sub-divisions.")
                for x in f_new:
                    if x['level'] == 'warning':
                        st.warning(f"**{x['check']}** — {x['detail']}")
                save_default = st.checkbox("Also save as this app's default Master", value=False,
                                           help="The default is used by anyone who has not uploaded their own — "
                                                "tick this only if it should replace the default for everyone. "
                                                "On Streamlit Community Cloud it is lost on redeploy/restart.")
                if st.button("Use this Master", type="primary"):
                    st.session_state.session_master = {'bytes': data, 'name': new_master.name}
                    if save_default:
                        save_master_file(new_master)
                    st.session_state.generated_buf = None
                    st.rerun()
        if st.session_state.get('session_master') and load_persisted_master():
            if st.button("Switch back to the app's default Master"):
                st.session_state.pop('session_master', None)
                st.session_state.generated_buf = None
                st.rerun()
        st.markdown("---")
        st.caption("Optional: replace the India Post logo shown in the header above.")
        new_logo = st.file_uploader("Header logo (.png/.jpg/.svg)", type=["png", "jpg", "jpeg", "svg"], key="logo_upload")
        if new_logo is not None and st.button("Save as header logo"):
            save_logo_file(new_logo)
            st.success("Logo updated.")
            st.rerun()

if master is None:
    st.stop()
if any(x['level'] == 'warning' for x in master_findings):
    for x in master_findings:
        if x['level'] == 'warning':
            st.warning(f"**{x['check']}** — {x['detail']}")

master_raw = master
sub_divs = get_sub_divisions(master_raw)

with st.expander("Report settings"):
    record_prefix = st.text_input(
        "Office record-number prefix", value=default_record_prefix(division_name),
        key=f"rec_prefix_{division_name}",
        help="Used in the 'No. <prefix>/G/Mails/MMU/Performance/<date>' line on the report. "
             "The default is derived from the Division name — change it if your office uses a different code.")
    signature_lines = signature_inputs(division_name, 'mmu')

st.divider()

# ------------------------------------------------------- input file grid ---
st.subheader("📁 Input Files — dates are always taken from the calendar pickers below, never from filenames")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**1️⃣ Delivery Productivity — Range 1**")
    dp_r1_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_r1_f", label_visibility="collapsed")
    dp_r1_from = st.date_input("From", key="dp_r1_from", value=None)
    dp_r1_to = st.date_input("To", key="dp_r1_to", value=None)

    st.markdown("**2️⃣ Delivery Productivity — Range 2** *(only if your source system caps exports at ~15 days)*")
    dp_r2_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_r2_f", label_visibility="collapsed")
    dp_r2_from = st.date_input("From", key="dp_r2_from", value=None)
    dp_r2_to = st.date_input("To", key="dp_r2_to", value=None)

    st.markdown("**3️⃣ Delivery Productivity — Single Date**")
    dp_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_sd_f", label_visibility="collapsed")
    dp_sd_date = st.date_input("Single Date", key="dp_sd_date", value=None)

with c2:
    st.markdown("**4️⃣ DSS Usage — Range**")
    dss_r_file = st.file_uploader("Upload CSV", type=["csv"], key="dss_r_f", label_visibility="collapsed")
    dss_r_from = st.date_input("From", key="dss_r_from", value=None)
    dss_r_to = st.date_input("To", key="dss_r_to", value=None)

    st.markdown("**5️⃣ DSS Usage — Single Date**")
    dss_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="dss_sd_f", label_visibility="collapsed")
    dss_sd_date = st.date_input("Single Date", key="dss_sd_date", value=None)

    st.markdown("**6️⃣ Letter Box Clearance — Range**")
    lb_r_file = st.file_uploader("Upload CSV", type=["csv"], key="lb_r_f", label_visibility="collapsed")
    lb_r_from = st.date_input("From", key="lb_r_from", value=None)
    lb_r_to = st.date_input("To", key="lb_r_to", value=None)

with c3:
    st.markdown("**7️⃣ Letter Box Clearance — Single Date**")
    lb_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="lb_sd_f", label_visibility="collapsed")
    lb_sd_date = st.date_input("Single Date", key="lb_sd_date", value=None)

    st.markdown("**8️⃣ COD Digital Transaction — Range**")
    cod_r_file = st.file_uploader("Upload CSV", type=["csv"], key="cod_r_f", label_visibility="collapsed")
    cod_r_from = st.date_input("From", key="cod_r_from", value=None)
    cod_r_to = st.date_input("To", key="cod_r_to", value=None)

    st.markdown("**9️⃣ COD Digital Transaction — Single Date**")
    cod_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="cod_sd_f", label_visibility="collapsed")
    cod_sd_date = st.date_input("Single Date", key="cod_sd_date", value=None)

st.divider()

# ------------------------------------------------------------- IDC (optional) ---
st.subheader("📦 IDC Delivery files (optional) — adds an 'IDC Report' sheet to this same workbook")
ic1, ic2 = st.columns(2)
with ic1:
    st.markdown("**IDC DPR** (delivery performance, one row per product per beat per day)")
    st.warning("✅ Before exporting from the portal: tick **ALL products**, and generate the "
               "**From–To date range report only**. No single-date report is needed.")
    idc_dpr_file = st.file_uploader("IDC DPR CSV", type=["csv"], key="mmu_idc_dpr_f", label_visibility="collapsed")
with ic2:
    st.markdown("**IDC DSS** (PDMS / DSS invoiced counts per beat per day)")
    st.warning("✅ Before exporting from the portal: generate the **From–To date range report only**. "
               "No single-date report is needed.")
    idc_dss_file = st.file_uploader("IDC DSS CSV", type=["csv"], key="mmu_idc_dss_f", label_visibility="collapsed")
st.caption("The IDC sheet uses the Range and Single Date derived from the calendar pickers above — no extra dates. "
           "Leave both empty to build the workbook without it.")

st.divider()


# --------------------------------------------------------- derive dates ---
def first_not_none(*vals):
    for v in vals:
        if v is not None:
            return v
    return None


canonical_single_date = first_not_none(dp_sd_date, dss_sd_date, lb_sd_date, cod_sd_date)
range_froms = [d for d in (dp_r1_from, dss_r_from, lb_r_from, cod_r_from) if d is not None]
range_tos = [d for d in (dp_r2_to or dp_r1_to, dss_r_to, lb_r_to, cod_r_to) if d is not None]
canonical_range_start = min(range_froms) if range_froms else None
canonical_range_end = max(range_tos) if range_tos else None

sd_values = [d for d in (dp_sd_date, dss_sd_date, lb_sd_date, cod_sd_date) if d is not None]
if len(set(sd_values)) > 1:
    st.warning(f"⚠️ The 4 Single Date pickers don't all agree ({sorted(set(sd_values))}). "
                "Using the earliest-listed one available (Delivery Productivity's, if provided) "
                "for the report title and office record number. Double-check this is intended.")

st.subheader("📅 Report dates (derived from your calendar picks above)")
d1, d2, d3 = st.columns(3)
d1.metric("Single Date (for title/record no.)", canonical_single_date.strftime('%d.%m.%Y') if canonical_single_date else "not set")
d2.metric("Range Start", canonical_range_start.strftime('%d.%m.%Y') if canonical_range_start else "not set")
d3.metric("Range End", canonical_range_end.strftime('%d.%m.%Y') if canonical_range_end else "not set")

if canonical_single_date is None:
    st.error("At least one Single Date picker must be set (Delivery Productivity's is preferred) "
              "before a report can be generated — the office record number depends on it.")
    st.stop()
if canonical_range_start is None or canonical_range_end is None:
    st.error("Set the From/To dates for at least one Range file — the Range columns are labelled with them.")
    st.stop()

record_date = canonical_single_date + datetime.timedelta(days=1)
st.caption(f"Office record number will read **No. {record_prefix}/G/Mails/MMU/Performance/"
            f"{record_date.strftime('%d.%m.%y')}** (day after Single Date).")

st.divider()


# ----------------------------------------------------------- file status ---
def status(label, present):
    return f"✅ {label}" if present else f"⬜ {label} — *not provided*"


file_checklist = [
    ("Delivery Productivity — Range 1", dp_r1_file is not None),
    ("Delivery Productivity — Range 2", dp_r2_file is not None),
    ("Delivery Productivity — Single Date", dp_sd_file is not None),
    ("DSS Usage — Range", dss_r_file is not None),
    ("DSS Usage — Single Date", dss_sd_file is not None),
    ("Letter Box Clearance — Range", lb_r_file is not None),
    ("Letter Box Clearance — Single Date", lb_sd_file is not None),
    ("COD Digital Transaction — Range", cod_r_file is not None),
    ("COD Digital Transaction — Single Date", cod_sd_file is not None),
]
missing = [label for label, present in file_checklist if not present]

st.subheader("✅ Input file status")
cols = st.columns(3)
for i, (label, present) in enumerate(file_checklist):
    cols[i % 3].markdown(status(label, present))

# ------------------------------------------------------------ read files ---
uploads = {'dp_r1': dp_r1_file, 'dp_r2': dp_r2_file, 'dp_sd': dp_sd_file, 'dss_r': dss_r_file,
           'dss_sd': dss_sd_file, 'lb_r': lb_r_file, 'lb_sd': lb_sd_file, 'cod_r': cod_r_file,
           'cod_sd': cod_sd_file}
slot_names = dict(zip(uploads, [x[0] for x in file_checklist]))
frames, read_errors = {}, []
for key, up in uploads.items():
    if up is None:
        frames[key] = None
        continue
    try:
        frames[key] = read_csv_any(up)
    except Exception as e:
        frames[key] = None
        read_errors.append({'level': 'error', 'check': f'{slot_names[key]}: file readable',
                            'detail': f'"{up.name}" could not be read as a CSV file ({e}).'})

# ------------------------------------------------------------ validation ---
st.divider()
st.subheader("🔎 Validation")
findings = master_findings + read_errors + run_validations(
    master_raw, master_info, frames, canonical_single_date, canonical_range_start, canonical_range_end)

# IDC files (optional): read + validate with the same canonical dates; errors block like any other file
idc_dpr_raw = idc_dss_raw = None
if idc_dpr_file is not None or idc_dss_file is not None:
    for _lbl, _up in (("IDC DPR", idc_dpr_file), ("IDC DSS", idc_dss_file)):
        if _up is None:
            continue
        try:
            _df = read_csv_any(_up)
            if _lbl == "IDC DPR":
                idc_dpr_raw = _df
            else:
                idc_dss_raw = _df
        except Exception as e:
            findings.append({'level': 'error', 'check': f'{_lbl}: file readable',
                              'detail': f'"{_up.name}" could not be read as a CSV file ({e}).'})
    findings += [dict(x, check='IDC — ' + x['check']) for x in idc.run_validations(
        idc_dpr_raw, idc_dss_raw, canonical_single_date, canonical_range_start, canonical_range_end)] \
        if idc_dpr_raw is not None or idc_dss_raw is not None else []
if has_errors(findings):
    show_findings(findings)
    st.error("Fix the error(s) above and re-upload — the report can't be built until they're resolved.")
    st.stop()

try:
    df, excl_r, excl_sd, avail = build_full(
        master_raw, frames['dp_r1'], frames['dp_r2'], frames['dp_sd'], frames['dss_r'], frames['dss_sd'],
        frames['cod_r'], frames['cod_sd'], frames['lb_r'], frames['lb_sd'])
except Exception as e:
    st.error(f"Couldn't process the uploaded files — check they match the expected column layout "
              f"for their category. Error: {e}")
    st.stop()
findings += check_percentage_bounds(df)
show_findings(findings)

idc_inputs, idc_sig = None, None
if idc_dpr_raw is not None:
    _d, _s = idc.prepare(idc_dpr_raw, idc_dss_raw)
    _offices = idc.offices_in(_d, _s)
    with st.expander("IDC report settings", expanded=len(_offices) > 1):
        _office = _offices[0]
        if len(_offices) > 1:
            _office = st.selectbox("IDC office (one IDC sheet per workbook)", _offices, key="mmu_idc_office")
        idc_label = st.text_input("Office name in the IDC title ('IDC <name> Delivery % …')",
                                   value=idc.office_short(_office), key=f"mmu_idc_label_{_office}")
    _dd, _ss = idc.filter_office(_d, _office), idc.filter_office(_s, _office)
    if _ss is not None and len(_ss) == 0:
        _ss = None
        st.warning(f"The IDC DSS file has no rows for '{_office}' — DSS Usage % will be blank on the IDC sheet.")
    if len(_dd) == 0:
        st.error(f"The IDC DPR file has no rows for '{_office}'.")
        st.stop()
    idc_inputs = {'dpr': _dd, 'dss': _ss, 'office_label': idc_label}
    idc_sig = (idc_dpr_file.name, idc_dpr_file.size, idc_dss_file.name if idc_dss_file else None,
               idc_dss_file.size if idc_dss_file else None, _office, idc_label)

st.divider()

# ------------------------------------------------------- live preview ---
total_label = f"{division_name or 'Division'} Total"
st.subheader("📊 Divisional Summary — live preview (Including B.Os, Range + Single Date)")
with st.container(border=True):
    ds = build_divisional_summary(df, total_label=total_label)
    preview_rows = ds['incl_bo']
    preview_df = pd.DataFrame(preview_rows)[
        ['sub_division', 'delivery_pct_r', 'delivery_pct_sd', 'dss_pct_r', 'dss_pct_sd',
         'cod_pct_r', 'cod_pct_sd', 'lb_pct_r', 'lb_pct_sd']
    ].rename(columns={
        'sub_division': 'Sub Division', 'delivery_pct_r': 'Delivery % (R)', 'delivery_pct_sd': 'Delivery % (SD)',
        'dss_pct_r': 'DSS % (R)', 'dss_pct_sd': 'DSS % (SD)', 'cod_pct_r': 'COD % (R)', 'cod_pct_sd': 'COD % (SD)',
        'lb_pct_r': 'LB % (R)', 'lb_pct_sd': 'LB % (SD)',
    })
    pct_cols = [c for c in preview_df.columns if c != 'Sub Division']
    st.dataframe(
        preview_df.style.format({c: lambda v: '' if pd.isna(v) else f'{v*100:.2f}%' for c in pct_cols}),
        use_container_width=True, hide_index=True,
    )
    total_row = ds['incl_bo'][-1]

    def fmt(v):
        return '' if pd.isna(v) else f"{v*100:.2f}%"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Delivery % (R) — Division", fmt(total_row['delivery_pct_r']),
               help=f"Single Date: {fmt(total_row['delivery_pct_sd'])}")
    m2.metric("DSS % (R) — Division", fmt(total_row['dss_pct_r']),
               help=f"Single Date: {fmt(total_row['dss_pct_sd'])}")
    m3.metric("COD % (R) — Division", fmt(total_row['cod_pct_r']),
               help=f"Single Date: {fmt(total_row['cod_pct_sd'])}")
    m4.metric("LB % (R) — Division", fmt(total_row['lb_pct_r']),
               help=f"Single Date: {fmt(total_row['lb_pct_sd'])}")

st.divider()

# ------------------------------------------------------------ generate ---
st.subheader("🛠️ Generate Report")

# Signature of everything the report depends on — used to stop a stale download being shown
# after the inputs (files, dates, Master, prefix) have changed.
current_sig = (
    tuple((k, up.name, up.size) if up is not None else (k, None, None) for k, up in uploads.items()),
    tuple(str(d) for d in (dp_r1_from, dp_r1_to, dp_r2_from, dp_r2_to, dp_sd_date, dss_r_from, dss_r_to,
                           dss_sd_date, lb_r_from, lb_r_to, lb_sd_date, cod_r_from, cod_r_to, cod_sd_date)),
    hash(master_bytes), record_prefix, tuple(signature_lines), idc_sig,
)


def do_generate():
    date_meta = {
        'DP R1': f"{dp_r1_from} to {dp_r1_to}" if dp_r1_from else None,
        'DP R2': f"{dp_r2_from} to {dp_r2_to}" if dp_r2_from else None,
        'DP SD': dp_sd_date, 'DSS R': f"{dss_r_from} to {dss_r_to}" if dss_r_from else None,
        'DSS SD': dss_sd_date, 'LB R': f"{lb_r_from} to {lb_r_to}" if lb_r_from else None,
        'LB SD': lb_sd_date, 'COD R': f"{cod_r_from} to {cod_r_to}" if cod_r_from else None,
        'COD SD': cod_sd_date,
    }
    n_visible, n_hidden = 5 + len(sub_divs) + (1 if idc_inputs else 0), 11 + (2 if idc_inputs else 0)
    with st.spinner(f"Building all {n_visible + n_hidden} sheets..."):
        buf = build_workbook(
            master_raw, df, excl_r, excl_sd, avail,
            canonical_single_date, canonical_range_start, canonical_range_end,
            frames['dp_r1'], frames['dp_r2'], frames['dp_sd'], frames['dss_r'], frames['dss_sd'],
            frames['lb_r'], frames['lb_sd'], frames['cod_r'], frames['cod_sd'], date_meta=date_meta,
            division_name=division_name or 'Division', record_prefix=record_prefix, findings=findings,
            signature=signature_lines, idc_inputs=idc_inputs,
        )
    st.session_state.generated_buf = buf
    st.session_state.generated_sig = current_sig
    st.session_state.generated_counts = (n_visible, n_hidden)
    st.session_state.confirm_missing = False


if missing and not st.session_state.confirm_missing:
    if st.button("🛠️ Generate MMU Hub Report (.xlsx)", type="primary"):
        st.session_state.confirm_missing = True
        st.rerun()

if st.session_state.confirm_missing:
    st.warning(
        f"**{len(missing)} of 9 files not uploaded:** " + "; ".join(missing) + ".\n\n"
        "Each missing parameter will show as **blank** everywhere in the workbook "
        "(never 0%, N/A, or a dash), and defaulter/exceptional tables for that "
        "parameter will note data wasn't provided rather than falsely implying zero "
        "defaulters.\n\n**Still generate the report?**"
    )
    gc1, gc2 = st.columns([1, 4])
    if gc1.button("✅ Yes, generate anyway", type="primary"):
        do_generate()
        st.rerun()
    if gc2.button("Cancel"):
        st.session_state.confirm_missing = False
        st.rerun()

if not missing:
    if st.button("🛠️ Generate MMU Hub Report (.xlsx)", type="primary"):
        do_generate()
        st.rerun()

if st.session_state.generated_buf is not None:
    if st.session_state.get('generated_sig') == current_sig:
        n_visible, n_hidden = st.session_state.get('generated_counts', (0, 0))
        st.success(f"Workbook built — {n_visible} visible sheets + {n_hidden} hidden sheets, per Build Spec v10.")
        fname = f"MMU_Hub_Report_{canonical_single_date.strftime('%d.%m.%Y')}.xlsx"
        st.download_button("⬇️ Download " + fname, data=st.session_state.generated_buf, file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("The files, dates or Master have changed since the last report was built — "
                "click Generate again to refresh it.")

st.divider()
with st.expander("About this app"):
    st.markdown(
        "Builds the **MMU Hub** performance workbook for a Postal Division from its Office Master File "
        "and the daily exports: 9 named input placeholders each with their own calendar picker (dates are "
        "never inferred from filenames), validation of every file before anything is built, graceful "
        "per-parameter degradation with an explicit 'Generate anyway?' confirmation, and a live Divisional "
        "Summary preview. The generated workbook follows **Build Spec v10** — Raw Data carries live "
        "cross-sheet formulas; every other visible sheet is hardcoded pre-computed values from this same "
        "engine, so nothing can silently drift between sheets."
    )
