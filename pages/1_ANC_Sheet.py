#!/usr/bin/env python3
"""
ANC Sheet Generator - Generate ANC sheets for any date
"""

import streamlit as st
from pathlib import Path
from datetime import datetime, date
import sys

st.set_page_config(
    page_title="ANC Sheet - Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

st.title("ANC Sheet Generator")

# Date picker (default to today)
selected_date = st.date_input("Select date", value=date.today())

# Convert to datetime
target_datetime = datetime.combine(selected_date, datetime.min.time())

# Project directory
project_dir = Path(__file__).parent.parent

# Check if file already exists for this date
date_str = selected_date.strftime("%m %d %y")
pattern = f"AUTO_{date_str}*.docx"
existing_files = list(project_dir.glob(pattern))

col1, col2 = st.columns([1, 1])

with col1:
    generate_btn = st.button("Generate ANC Sheet", type="primary")

with col2:
    if existing_files:
        with open(existing_files[0], "rb") as f:
            st.download_button(
                "Download Existing",
                data=f.read(),
                file_name=existing_files[0].name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if existing_files:
    st.info(f"Existing file: **{existing_files[0].name}**")

if generate_btn:
    with st.spinner(f"Generating ANC sheet for {selected_date.strftime('%A, %B %d, %Y')}..."):
        try:
            # Add project dir to path for imports
            sys.path.insert(0, str(project_dir))
            from anc_generator import generate_anc_for_date

            output_path = generate_anc_for_date(
                target_datetime,
                output_dir=str(project_dir),
                output_format='docx',
                validate=True,
                notify_on_failure=False
            )

            if output_path:
                st.success(f"Generated: **{Path(output_path).name}**")

                # Offer download
                with open(output_path, "rb") as f:
                    st.download_button(
                        "Download",
                        data=f.read(),
                        file_name=Path(output_path).name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_new"
                    )
            else:
                st.error("Generation failed. Check logs for details.")

        except Exception as e:
            st.error(f"Error: {e}")
