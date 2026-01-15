#!/usr/bin/env python3
"""
Monday Shuffle - Redistribution Mode

Identify patients on wrong teams and suggest geographic reassignments.
"""

import streamlit as st
import re
from dataclasses import dataclass
from typing import Optional
from collections import defaultdict
from PIL import Image

try:
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# =============================================================================
# CONFIGURATION (same as main app)
# =============================================================================

FLOOR_TO_TEAMS = {
    '3W': [1, 2, 3], '3E': [1, 2, 3],
    '4W': [4, 11], '4E': [4, 11],
    '5W': [5, 10], '5E': [5, 10],
    '6W': [6, 12], '6E': [6, 12],
    '7W': [7, 9], '7E': [7, 9],
    '8E': [8, 13], '8W': [8, 13],
    'IMCU': [1, 2, 3],
    'BOYER': [12, 6],
}

ALL_TEAMS = list(range(1, 16))
OVERFLOW_TEAMS = [14, 15]
IMCU_TEAMS = [1, 2, 3]


@dataclass
class ExistingPatient:
    room: str
    current_team: int
    floor: str


def normalize_floor(location: str) -> Optional[str]:
    """Parse room number to floor."""
    original = location.strip().upper()

    if 'IMCU' in original:
        return 'IMCU'
    if 'OVERNIGHT' in original or 'ONR' in original or 'RECOVERY' in original:
        return 'BOYER'
    if any(x in original for x in ['ED', 'EMERGENCY', 'ER ']):
        return None
    if 'BOYER' in original:
        return 'BOYER'

    location = re.sub(r'[A-Z]$', '', original)

    match = re.search(r'(\d+)\s*([EW]|EAST|WEST)', location)
    if match:
        floor_num = match.group(1)
        direction = match.group(2)[0]
        if floor_num == '9' and direction == 'W':
            return 'BOYER'
        return f"{floor_num}{direction}"

    match = re.search(r'\b(\d)(\d{2})\b', location)
    if match:
        floor_num = match.group(1)
        room_num = int(match.group(2))

        if floor_num == '7' and room_num >= 50:
            return 'BOYER'
        if floor_num == '8' and room_num >= 50:
            return 'BOYER'
        if floor_num == '9':
            return 'BOYER'

        if 1 <= room_num <= 20:
            return f"{floor_num}W"
        elif 30 <= room_num < 50:
            return f"{floor_num}E"
        else:
            return f"{floor_num}?"

    return None


def get_geographic_teams(floor: str) -> list[int]:
    """Get teams that cover a given floor."""
    if not floor:
        return []
    if floor in FLOOR_TO_TEAMS:
        return FLOOR_TO_TEAMS[floor]
    if floor.endswith('?'):
        floor_num = floor[:-1]
        east_teams = FLOOR_TO_TEAMS.get(f"{floor_num}E", [])
        west_teams = FLOOR_TO_TEAMS.get(f"{floor_num}W", [])
        return list(set(east_teams + west_teams))
    return []


def analyze_patients(
    patients: list[ExistingPatient],
    closed_teams: set[int] = None
) -> tuple[list[tuple[ExistingPatient, list[int]]], list[ExistingPatient]]:
    """Analyze patients and separate into wrong team vs OK."""
    closed_teams = closed_teams or set()

    wrong_team = []
    ok_team = []

    for patient in patients:
        geo_teams = get_geographic_teams(patient.floor) if patient.floor else []
        acceptable_teams = [t for t in geo_teams if t not in closed_teams]

        if patient.current_team in acceptable_teams:
            ok_team.append(patient)
        elif acceptable_teams:
            wrong_team.append((patient, acceptable_teams))
        else:
            wrong_team.append((patient, []))

    return wrong_team, ok_team


def extract_from_ocr(text: str) -> list[tuple[str, int]]:
    """
    Extract room-team pairs from OCR text.

    Tries multiple strategies:
    1. Same-line matching (room and Med X on same line)
    2. Column matching (extract all rooms, all teams, zip them)
    """
    pairs = []
    seen_rooms = set()

    # Room patterns: 3-digit rooms (304A), special rooms (MAIN, RZ08, YZ26)
    room_pattern = r'\b(\d{3}[A-Z]?|[A-Z]{2,4}\d{0,2}[A-Z]?|\d{4}[A-Z]?)\b'

    # Strategy 1: Look for room and team on same line
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue

        # Skip header lines
        if 'Primary' in line or 'Bed' in line and 'Team' in line:
            continue

        room_match = re.search(room_pattern, line, re.IGNORECASE)
        team_match = re.search(r'Med\s*(\d{1,2})', line, re.IGNORECASE)

        if room_match and team_match:
            room = room_match.group(1).upper()
            team = int(team_match.group(1))
            # Skip if room looks like a header or invalid
            if room in ('BED', 'PRIMARY', 'TEAM'):
                continue
            if room not in seen_rooms and 1 <= team <= 15:
                seen_rooms.add(room)
                pairs.append((room, team))

    # Strategy 2: Column matching - always try this to catch more
    # More flexible room pattern for column extraction
    all_rooms = re.findall(r'\b(\d{3,4}[A-Z]?)\b', text, re.IGNORECASE)
    # Also find special rooms like MAIN, RZ08, YZ26
    special_rooms = re.findall(r'\b([A-Z]{2,4}\d{1,2})\b', text)
    special_rooms += re.findall(r'\b(MAIN|ICU\d*|IMCU\d*)\b', text, re.IGNORECASE)

    unique_rooms = []
    for r in all_rooms + special_rooms:
        r_upper = r.upper()
        if r_upper not in unique_rooms and r_upper not in ('BED', 'PRIMARY', 'TEAM', 'MED'):
            unique_rooms.append(r_upper)

    all_teams = re.findall(r'Med\s*(\d{1,2})', text, re.IGNORECASE)
    all_teams = [int(t) for t in all_teams if 1 <= int(t) <= 15]

    # If strategy 1 found enough and counts are close, return
    if len(pairs) >= 5 and abs(len(pairs) - len(all_teams)) <= 3:
        return pairs

    # Strategy 2: Column matching - be aggressive
    if len(unique_rooms) > 0 and len(all_teams) > 0:
        # Always try to match what we can
        min_len = min(len(unique_rooms), len(all_teams))
        for room, team in zip(unique_rooms[:min_len], all_teams[:min_len]):
            if room not in seen_rooms:
                seen_rooms.add(room)
                pairs.append((room, team))

    return pairs


# =============================================================================
# STREAMLIT UI
# =============================================================================

st.set_page_config(
    page_title="Monday Shuffle - Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

st.title("Monday Shuffle")
st.markdown("Identify patients on wrong teams for manual reassignment in Epic.")

# Closed teams input
closed_input = st.text_input(
    "Closed teams (comma-separated, e.g., '14, 15')",
    value="14, 15",
    help="Enter team numbers that are closed"
)
closed_teams = set()
if closed_input:
    for t in closed_input.split(','):
        t = t.strip()
        if t.isdigit():
            closed_teams.add(int(t))

st.markdown("---")

# Input method tabs
tab1, tab2 = st.tabs(["Screenshot OCR", "Manual Entry"])

# Use session state to persist patients across reruns
if 'all_patients' not in st.session_state:
    st.session_state.all_patients = []

with tab1:
    st.markdown("""
    **Upload Epic screenshots** showing room numbers and team assignments.

    For best results, crop screenshots to show just the Room and Team columns.
    """)

    if not OCR_AVAILABLE:
        st.error("OCR (pytesseract) not available. Use Manual Entry tab instead.")
    else:
        uploaded_files = st.file_uploader(
            "Upload Epic screenshots",
            type=['png', 'jpg', 'jpeg'],
            accept_multiple_files=True
        )

        if uploaded_files:
            if st.button("Process Screenshots", type="primary", key="ocr_btn"):
                all_pairs = []

                for uploaded_file in uploaded_files:
                    image = Image.open(uploaded_file)
                    gray_image = image.convert('L')

                    raw_text = pytesseract.image_to_string(gray_image)

                    with st.expander(f"Raw OCR from {uploaded_file.name}"):
                        st.code(raw_text)
                        # Debug info
                        rooms_found = re.findall(r'\b(\d{3,4}[A-Z]?)\b', raw_text, re.IGNORECASE)
                        teams_found = re.findall(r'Med\s*(\d{1,2})', raw_text, re.IGNORECASE)
                        st.caption(f"Debug: {len(rooms_found)} room patterns, {len(teams_found)} Med patterns")

                    pairs = extract_from_ocr(raw_text)
                    all_pairs.extend(pairs)

                    st.info(f"Found {len(pairs)} room-team pairs in {uploaded_file.name}")

                # Deduplicate and store in session state
                seen_rooms = set()
                st.session_state.all_patients = []
                for room, team in all_pairs:
                    if room not in seen_rooms:
                        seen_rooms.add(room)
                        floor = normalize_floor(room)
                        st.session_state.all_patients.append(
                            ExistingPatient(room=room, current_team=team, floor=floor)
                        )

                if st.session_state.all_patients:
                    st.success(f"Total: {len(st.session_state.all_patients)} unique patients extracted")

                    with st.expander("Extracted data (verify this is correct)"):
                        extracted_text = ""
                        for p in sorted(st.session_state.all_patients, key=lambda x: x.room):
                            extracted_text += f"{p.room} Med {p.current_team}\n"
                        st.code(extracted_text)
                else:
                    st.warning("No room-team pairs found. Check the raw OCR output above.")
                    st.markdown("""
                    **Troubleshooting:**
                    - Make sure the screenshot shows both Room and Team columns
                    - Try cropping to just those two columns
                    - Use Manual Entry tab if OCR isn't working
                    """)

with tab2:
    st.markdown("""
    **Enter room and current team** - one per line.

    Format: `Room Team` (e.g., `304A 1` or `304A Med 1`)
    """)

    paste_input = st.text_area(
        "Room and Team (one per line)",
        height=400,
        placeholder="304A 1\n343B 5\n534 Med 10\n435A 7\n..."
    )

    if st.button("Analyze Teams", type="primary", key="manual_btn"):
        if not paste_input:
            st.warning("Please enter room and team data first.")
        else:
            seen_rooms = set()
            parse_errors = []
            st.session_state.all_patients = []

            for line in paste_input.strip().split('\n'):
                line = line.strip()
                if not line:
                    continue

                room_match = re.search(r'\b(\d{3}[A-Z]?)\b', line, re.IGNORECASE)
                if not room_match:
                    parse_errors.append(f"No room found: {line}")
                    continue

                room = room_match.group(1).upper()

                team_match = re.search(r'(?:Med\s*)?(\d{1,2})\b', line[room_match.end():], re.IGNORECASE)
                if not team_match:
                    parse_errors.append(f"No team found: {line}")
                    continue

                team_num = int(team_match.group(1))

                if team_num < 1 or team_num > 15:
                    parse_errors.append(f"Invalid team {team_num}: {line}")
                    continue

                if room in seen_rooms:
                    continue
                seen_rooms.add(room)

                floor = normalize_floor(room)
                st.session_state.all_patients.append(
                    ExistingPatient(room=room, current_team=team_num, floor=floor)
                )

            if parse_errors:
                with st.expander(f"Parse warnings ({len(parse_errors)})"):
                    for err in parse_errors:
                        st.warning(err)

            st.success(f"Parsed {len(st.session_state.all_patients)} patients")

# Process results if we have patients
if st.session_state.all_patients:
    all_patients = st.session_state.all_patients
    wrong_team, ok_team = analyze_patients(all_patients, closed_teams)

    st.markdown("---")
    st.subheader("Analysis Results")

    # Metrics row: Total / Need Reassignment / Team Correct
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Patients", len(all_patients))
    col2.metric("Need Reassignment", len(wrong_team))
    col3.metric("Team Correct", len(ok_team))

    # Calculate current census
    team_census = defaultdict(int)
    for patient in all_patients:
        team_census[patient.current_team] += 1

    # Generate recommendations - 100% geographic (use current census as tie-breaker only)
    recommendations = []

    for patient, acceptable in sorted(wrong_team, key=lambda x: x[0].room):
        if acceptable:
            # Pick team with lowest current census from geographic options
            best_team = min(acceptable, key=lambda t: team_census.get(t, 0))
            recommendations.append((patient, best_team))
        else:
            recommendations.append((patient, None))

    # Calculate projected census if all recommendations followed
    projected_census = dict(team_census)
    for patient, rec_team in recommendations:
        if rec_team:
            projected_census[rec_team] = projected_census.get(rec_team, 0) + 1
            projected_census[patient.current_team] = projected_census.get(patient.current_team, 0) - 1

    # Three columns: Census, Needs Reassignment, Team Correct
    res_col1, res_col2, res_col3 = st.columns(3)

    with res_col1:
        st.markdown("### Census")
        census_text = "Team  Now  +/-  =New\n"
        census_text += "-" * 22 + "\n"
        for team in ALL_TEAMS:
            if team in closed_teams:
                census_text += f"Med {team:2d}   CLOSED\n"
            else:
                current = team_census.get(team, 0)
                projected = projected_census.get(team, 0)
                change = projected - current
                imcu = "*" if team in IMCU_TEAMS else " "
                if change > 0:
                    census_text += f"Med {team:2d}{imcu} {current:2d}  +{change:2d}  ={projected:2d}\n"
                elif change < 0:
                    census_text += f"Med {team:2d}{imcu} {current:2d}  {change:3d}  ={projected:2d}\n"
                else:
                    census_text += f"Med {team:2d}{imcu} {current:2d}    0  ={projected:2d}\n"
        census_text += "\n* = IMCU"
        st.code(census_text, language=None)

    with res_col2:
        st.markdown("### Needs Reassignment")
        if recommendations:
            wrong_text = ""
            for patient, rec_team in recommendations:
                # Format: "Room  Med XX -> Med YY" with aligned arrows
                room_padded = f"{patient.room:5}"
                current_padded = f"Med {patient.current_team:2d}"
                if rec_team:
                    new_padded = f"Med {rec_team:2d}"
                    wrong_text += f"{room_padded} {current_padded} -> {new_padded}\n"
                else:
                    wrong_text += f"{room_padded} {current_padded} -> ?\n"
            st.code(wrong_text, language=None)
        else:
            st.info("All patients on correct teams!")

    with res_col3:
        st.markdown("### Team Correct")
        if ok_team:
            ok_text = ""
            for patient in sorted(ok_team, key=lambda x: x.room):
                ok_text += f"{patient.room:5} Med {patient.current_team:2d}\n"
            st.code(ok_text, language=None)
        else:
            st.info("None yet")
