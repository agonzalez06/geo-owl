#!/usr/bin/env python3
"""
Geo Owl - Main Entry Point
Uses st.navigation for custom sidebar labels
"""

import streamlit as st

st.set_page_config(
    page_title="Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

# Define pages with custom titles
geo_owl_page = st.Page(
    "geo_placer_web.py",
    title="Geo Owl",
    icon=":material/home:",
    default=True
)

monday_shuffle_page = st.Page(
    "pages/2_Monday_Shuffle.py",
    title="Monday Shuffle",
    icon=":material/shuffle:"
)

# Set up navigation
pg = st.navigation([geo_owl_page, monday_shuffle_page])
pg.run()
