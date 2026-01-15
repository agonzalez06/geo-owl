#!/usr/bin/env python3
"""
Geo Owl - Main Entry Point

Redirects to the Overnight Redis page.
"""

import streamlit as st

st.set_page_config(
    page_title="Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

# Redirect to Overnight Redis page
st.switch_page("pages/1_Overnight_Redis.py")
