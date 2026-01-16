#!/usr/bin/env python3
"""
ANC Sheet Viewer - View the current ANC sheet
"""

import streamlit as st
import os
from pathlib import Path
from datetime import datetime
import glob

st.set_page_config(
    page_title="ANC Sheet - Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

st.title("ANC Sheet")
st.markdown("View and download the current ANC sheet.")

# Find ANC sheet files in the project directory
project_dir = Path(__file__).parent.parent
anc_files = list(project_dir.glob("AUTO_*.docx"))

# Sort by modification time (most recent first)
anc_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)

if anc_files:
    st.subheader("Available ANC Sheets")

    for anc_file in anc_files[:5]:  # Show last 5
        mod_time = datetime.fromtimestamp(anc_file.stat().st_mtime)
        file_size = anc_file.stat().st_size / 1024  # KB

        col1, col2, col3 = st.columns([3, 1, 1])

        with col1:
            st.markdown(f"**{anc_file.name}**")
        with col2:
            st.caption(f"{mod_time.strftime('%m/%d %H:%M')}")
        with col3:
            with open(anc_file, "rb") as f:
                st.download_button(
                    "Download",
                    data=f.read(),
                    file_name=anc_file.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{anc_file.name}"
                )

    # Preview of most recent
    st.markdown("---")
    st.subheader("Preview")
    st.info("Download the DOCX file to view the full ANC sheet. Preview coming soon.")

else:
    st.warning("No ANC sheets found. Generate one using the ANC generator.")
    st.markdown("""
    Run the ANC generator from terminal:
    ```
    python anc_generator.py
    ```
    """)
