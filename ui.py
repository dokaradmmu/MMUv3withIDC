"""Shared UI pieces for the MMU Hub app: branded header, Master lookup (session upload wins over saved default)."""
import base64
import html
import io
import mimetypes

import streamlit as st

from signature import default_signature
from engine import read_master_full, load_persisted_master, load_persisted_logo

INDIA_POST_LOGO = "https://upload.wikimedia.org/wikipedia/en/3/32/India_Post.svg"
NATIONAL_EMBLEM = "https://upload.wikimedia.org/wikipedia/commons/5/55/Emblem_of_India.svg"
MAROON = "#8B0000"
MAROON_DARK = "#5C0000"
GOLD = "#F5B700"


@st.cache_data(show_spinner=False)
def parse_master(data: bytes):
    """Validate + parse a Master file (cached on its content)."""
    return read_master_full(io.BytesIO(data))


def logo_src():
    """The header logo as something a browser can load. A persisted logo lives on the server's
    disk, so it is embedded as a data: URI (a bare file path cannot be loaded by the browser)."""
    path = load_persisted_logo()
    if path:
        try:
            mime = mimetypes.guess_type(path)[0] or 'image/png'
            with open(path, 'rb') as fh:
                return f"data:{mime};base64,{base64.b64encode(fh.read()).decode()}"
        except OSError:
            pass
    return INDIA_POST_LOGO


# ----------------------------------------------------------- master source ---
def active_master_bytes():
    """Session upload wins; else the app's saved default. Returns (bytes, source_label) or (None, None)."""
    sess = st.session_state.get('session_master')
    if sess:
        return sess['bytes'], f"your uploaded Master ({sess['name']}) — this session only"
    p = load_persisted_master()
    if p:
        with open(p, 'rb') as fh:
            return fh.read(), "this app's saved default Master"
    return None, None



def _flat(s):
    """Strip leading indentation (Markdown treats 4+ leading spaces as a code block)."""
    return "\n".join(l.strip() for l in s.split("\n"))


def load_master():
    """-> (master_df or None, info dict, findings, source label)."""
    mb, src = active_master_bytes()
    if not mb:
        return None, {}, [], None
    m, info, f = parse_master(mb)
    return m, info, f, src


def header_text(master, info):
    division_name = info.get('division_name', '') if master is not None else ''
    short = division_name[:-len(' Division')] if division_name.lower().endswith(' division') else division_name
    title = f"{short} Postal Division MMU Hub" if short else "MMU Hub Report"
    place = ", ".join(x for x in (info.get('region_name', ''), info.get('circle_name', '')) if x)
    sub = (place + " — " if place else "") + "Mail Motor / Mobile Unit Performance Monitoring"
    return title, sub


def render_header(header_title, header_sub):
    # ------------------------------------------------------------------- CSS ---
    st.markdown(_flat(f"""
    <style>
        .gov-topbar {{
            background:#f2f2f2; padding:4px 16px; font-size:12px; color:#444;
            border-bottom:1px solid #ddd; margin:-1rem -1rem 0 -1rem;
        }}
        .gov-header {{
            background: linear-gradient(135deg, {MAROON} 0%, {MAROON_DARK} 100%);
            padding: 18px 28px; margin: 0 -1rem 1.2rem -1rem;
            border-bottom: 4px solid {GOLD};
            display:flex; align-items:center; gap:18px;
        }}
        .gov-header img {{ height:54px; background:white; border-radius:6px; padding:4px; }}
        .gov-header .titles h1 {{
            color:white; font-size:26px; font-weight:800; margin:0; line-height:1.25;
            font-family: 'Segoe UI', Arial, sans-serif;
        }}
        .gov-header .titles p {{
            color:#f0e0e0; font-size:13px; margin:2px 0 0 0; font-family: 'Nirmala UI', Arial, sans-serif;
        }}
        div[data-testid="stMetricValue"] {{ color:{MAROON_DARK}; }}
        .stButton>button[kind="primary"] {{ background-color:{MAROON}; border-color:{MAROON}; }}
    </style>
    <div class="gov-topbar">🇮🇳 Government of India &nbsp;|&nbsp; Ministry of Communications &nbsp;|&nbsp; Department of Posts</div>
    """), unsafe_allow_html=True)

    st.markdown(_flat(f"""
    <div class="gov-header">
        <img src="{NATIONAL_EMBLEM}" />
        <img src="{logo_src()}" />
        <div class="titles">
            <h1>{html.escape(header_title)}</h1>
            <p>भारत सरकार | Government of India · संचार मंत्रालय · डाक विभाग · Department of Posts</p>
            <p>{html.escape(header_sub)}</p>
        </div>
    </div>
    """), unsafe_allow_html=True)


def signature_inputs(division_name, key):
    """Four editable lines for the signature block printed under every report sheet. Blank lines are skipped."""
    defaults = default_signature(division_name)
    st.markdown("**Signature block** — printed at the bottom of every report sheet (leave a line blank to skip it)")
    return [st.text_input(f"Signature line {i + 1}", value=defaults[i] if i < len(defaults) else "",
                          key=f"{key}_sig{i}_{division_name}") for i in range(4)]
