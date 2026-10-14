"""
MMU Hub — Streamlit webapp (Department of Posts).

Two reports, chosen from the sidebar:
  * MMU Daily Report   (mmu_report.py)  — Master-driven 4-KPI report for every office of a Division
  * IDC Delivery Report (idc_report.py) — per-delivery-person Delivery / DSS / Parcel / RTS / Redirected report

Run locally:   streamlit run app.py
Deploy:        Streamlit Community Cloud (main file = app.py).
"""
import streamlit as st

st.set_page_config(page_title="MMU Hub", page_icon="📮", layout="wide")

pages = [
    st.Page("mmu_report.py", title="MMU Daily Report", icon="📮", default=True),
    st.Page("idc_report.py", title="IDC Delivery Report", icon="📦", url_path="idc"),
]
st.navigation(pages).run()
