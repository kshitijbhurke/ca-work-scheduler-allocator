"""Shared look of the two web apps, matching the desktop colours."""

import streamlit as st

NAVY = "#0F2A47"
OFF_DAY_BG = "#FBE3D6"
TODAY_BG = "#E0EAFF"
ALLOCATED_BG = "#DBEAFE"
AMBER_BG = "#FEF3C7"
AMBER_TEXT = "#92400E"
WIP_BG = "#FFF6E0"
COMPLETE_BG = "#E8F8EE"
OVERDUE_TEXT = "#B91C1C"
TEXT = "#0F172A"

_CSS = f"""
<style>
.block-container {{ padding-top: 3.5rem; max-width: 100%; }}
.app-header {{ background: {NAVY}; color: white; padding: 1rem 1.5rem; border-radius: 8px;
               border-bottom: 4px solid #2563EB; margin-bottom: 1rem; }}
.app-header h1 {{ color: white; font-size: 1.7rem; margin: 0; padding: 0; }}
.app-header p {{ color: #A9C1DE; margin: 0.1rem 0 0; font-size: 0.9rem; }}
.legend span {{ display: inline-block; padding: 0.1rem 0.6rem; margin-right: 0.4rem;
                border-radius: 4px; font-size: 0.8rem; color: {TEXT}; }}
</style>
"""


def header(title, subtitle):
    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown(f'<div class="app-header"><h1>{title}</h1><p>{subtitle}</p></div>',
                unsafe_allow_html=True)


def legend(*items):
    """A row of colour keys: (label, background, text colour)."""
    spans = "".join(f'<span style="background:{bg};color:{fg}">{label}</span>'
                    for label, bg, fg in items)
    st.markdown(f'<div class="legend">{spans}</div>', unsafe_allow_html=True)
