"""IDC Delivery Report page — one row per delivery person (postman)."""
import datetime

import pandas as pd
import streamlit as st

import idc
from engine import read_csv_any, has_errors, default_record_prefix
from ui import load_master, header_text, render_header, signature_inputs

master, master_info, master_findings, master_source = load_master()
division_name = master_info.get('division_name', '') if master is not None else ''
title, sub = header_text(master, master_info)
render_header("IDC Delivery Report" if not division_name else f"{title.replace(' MMU Hub', '')} — IDC Delivery Report",
              "Delivery %, DSS Usage %, Parcel Delivery %, RTS % and Redirected % per delivery person")

st.session_state.setdefault('idc_buf', None)
st.session_state.setdefault('idc_sig', None)


def show_findings(findings):
    errs = [x for x in findings if x['level'] == 'error']
    warns = [x for x in findings if x['level'] == 'warning']
    infos = [x for x in findings if x['level'] == 'info']
    oks = [x for x in findings if x['level'] == 'pass']
    for x in errs:
        st.error(f"**{x['check']}** — {x['detail']}")
    for x in warns:
        st.warning(f"**{x['check']}** — {x['detail']}")
    for x in infos:
        st.info(f"**{x['check']}** — {x['detail']}")
    if oks:
        with st.expander(f"✅ {len(oks)} check(s) passed"):
            for x in oks:
                st.write(f"✔ **{x['check']}** — {x['detail']}")


# ----------------------------------------------------------------- inputs ---
st.subheader("📁 Input files — dates come only from the calendar pickers, never from filenames")
c1, c2 = st.columns(2)
with c1:
    st.markdown("**IDC DPR** (delivery performance, one row per product per beat per day)")
    st.warning("✅ Tick **ALL products**; generate the **From–To date range report only** (no single-date report).")
    dpr_file = st.file_uploader("IDC DPR CSV", type=["csv"], key="idc_dpr_f", label_visibility="collapsed")
with c2:
    st.markdown("**IDC DSS** (PDMS / DSS invoiced counts per beat per day) — optional")
    st.warning("✅ Generate the **From–To date range report only** (no single-date report).")
    dss_file = st.file_uploader("IDC DSS CSV", type=["csv"], key="idc_dss_f", label_visibility="collapsed")

d1, d2, d3 = st.columns(3)
with d1:
    r_from = st.date_input("Range From (e.g. 1st of the month)", key="idc_from", value=None)
with d2:
    r_to = st.date_input("Range To (e.g. yesterday)", key="idc_to", value=None)
with d3:
    single = st.date_input("Single Date (e.g. yesterday)", key="idc_single", value=None)
st.caption("The Range columns use every row of the uploaded files; the Single Date columns use only that day's rows.")

dpr_raw = dss_raw = None
read_errors = []
for label, f in (("IDC DPR", dpr_file), ("IDC DSS", dss_file)):
    if f is None:
        continue
    try:
        df = read_csv_any(f)
        if label == "IDC DPR":
            dpr_raw = df
        else:
            dss_raw = df
    except Exception as e:      # unreadable / not a CSV
        read_errors.append(f"Could not read the {label} file: {e}")
for m in read_errors:
    st.error(m)

if dpr_file is None and not read_errors:
    st.info("Upload the IDC DPR CSV to begin.")
    st.stop()
if dpr_raw is None:
    st.stop()

st.divider()
st.subheader("🔎 Validation")
findings = idc.run_validations(dpr_raw, dss_raw, single, r_from, r_to)
show_findings(findings)
if has_errors(findings):
    st.error("Fix the errors above to continue.")
    st.stop()

dpr_all, dss_all = idc.prepare(dpr_raw, dss_raw)
offices = idc.offices_in(dpr_all, dss_all)

with st.expander("Report settings", expanded=len(offices) > 1):
    office = offices[0]
    if len(offices) > 1:
        office = st.selectbox("Office (one report is built per office)", offices, key="idc_office")
    office_label = st.text_input("Office name in the title ('IDC <name> Delivery % …')",
                                 value=idc.office_short(office), key=f"idc_label_{office}")
    record_prefix = st.text_input(
        "Office record-number prefix", value=default_record_prefix(division_name) if division_name else "KRDDN",
        key=f"idc_prefix_{division_name}",
        help="Used in the 'No. <prefix>/G/MMU/IDC/Delivery Performance/<Single Date>' line. "
             "The default is derived from the Master's Division name when a Master is loaded.")
    signature_lines = signature_inputs(division_name or 'Karad Division', 'idc')

dpr = idc.filter_office(dpr_all, office)
dss = idc.filter_office(dss_all, office)
if dss is not None and len(dss) == 0:
    dss = None
    st.warning(f"The DSS file has no rows for '{office}' — DSS Usage % columns will be blank.")
if len(dpr) == 0:
    st.error(f"The DPR file has no rows for '{office}'.")
    st.stop()

# ---------------------------------------------------------------- preview ---
T, TOT = idc.compute_table(dpr, dss, single)
st.divider()
st.subheader("📊 Preview — sorted by Delivery % (range), smallest first")
pct = lambda x: "" if pd.isna(x) else f"{x * 100:.2f}%"
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Delivery % (range)", pct(TOT['dl_r']) or "—")
m2.metric(f"Delivery % ({single:%d %b})", pct(TOT['dl_s']) or "—")
m3.metric("DSS Usage % (range)", pct(TOT['ds_r']) or "—")
m4.metric("Parcel Delivery %", pct(TOT['p_dl']) or "—")
m5.metric("RTS % / Redirected %", f"{pct(TOT['rts'])} / {pct(TOT['red'])}")
show = pd.DataFrame({
    "Sr. No": range(1, len(T) + 1), "Delivery Person": T.person,
    "Delivery % (range)": T.dl_r.map(pct), f"Delivery % ({single:%d-%b})": T.dl_s.map(pct),
    "DSS % (range)": T.ds_r.map(pct), f"DSS % ({single:%d-%b})": T.ds_s.map(pct),
    "Parcels Invoiced": T.p_inv.astype(int), "Parcel Delivery %": T.p_dl.map(pct),
    "RTS %": T.rts.map(pct), "Redirected %": T.red.map(pct)})
try:
    st.dataframe(show, hide_index=True, width="stretch")
except TypeError:                                   # older Streamlit
    st.dataframe(show, hide_index=True, use_container_width=True)
st.caption("The Excel file adds the colour formatting, the total row and the signature block.")

# --------------------------------------------------------------- generate ---
st.divider()
st.subheader("🛠️ Generate Report")
sig = (dpr_file.name, dpr_file.size, dss_file.name if dss_file else None, dss_file.size if dss_file else None,
       str(r_from), str(r_to), str(single), office, office_label, record_prefix, tuple(signature_lines))
if st.button("Generate IDC Report", type="primary", key="idc_gen"):
    with st.spinner("Building the workbook…"):
        buf, _, _ = idc.build_workbook(dpr, dss, office_label, record_prefix, single, r_from, r_to, findings,
                                     signature=signature_lines)
    st.session_state.idc_buf = buf.getvalue()
    st.session_state.idc_sig = sig
if st.session_state.idc_buf is not None:
    if st.session_state.idc_sig == sig:
        fname = f"IDC_Report_{single:%d.%m.%Y}.xlsx"
        st.success(f"Report ready — 1 visible sheet (IDC Report) + 3 hidden sheets (sources and validation log).")
        st.download_button("⬇️ Download " + fname, data=st.session_state.idc_buf, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="idc_dl")
    else:
        st.warning("Inputs changed since this report was generated — click **Generate IDC Report** again.")
