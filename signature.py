"""Signature block printed at the bottom of every report sheet (no Streamlit / Excel dependencies)."""

KARAD_LINES = ['डाक अधीक्षक, कराड डाक विभाग/',
               'Superintendent of Posts, Karad Division',
               'कराड- 415110/ Karad -415110 ',
               '4J2 6T7M FJP']


def default_signature(division_name):
    """Karad gets the full 4-line block; any other Division gets the English line only (editable in the UI)."""
    name = (division_name or '').strip()
    if 'karad' in name.lower():
        return list(KARAD_LINES)
    return [f'Superintendent of Posts, {name}'] if name else []


def clean_lines(lines):
    """Drop blank lines, keep order."""
    return [l for l in (lines or []) if str(l).strip()]
