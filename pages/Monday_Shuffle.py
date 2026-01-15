#!/usr/bin/env python3
"""
Monday Shuffle - Redistribution Mode

Upload Epic screenshots to extract patient data and optimize redistribution.
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
IMCU_CAP = 10
SOFT_CAP = 14


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


def parse_epic_screenshot(image) -> list[ExistingPatient]:
    """Parse Epic patient list screenshot using OCR."""
    if not OCR_AVAILABLE:
        return []

    text = pytesseract.image_to_string(image)

    patients = []
    seen_rooms = set()

    # First try: look for room and team on same line
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue

        team_match = re.search(r'Med\s*(\d+)', line, re.IGNORECASE)
        room_match = re.search(r'\b(\d{3}[A-Z]?)\b', line)

        if team_match and room_match:
            team_num = int(team_match.group(1))
            room = room_match.group(1)

            if room not in seen_rooms:
                seen_rooms.add(room)
                floor = normalize_floor(room)
                patients.append(ExistingPatient(room=room, current_team=team_num, floor=floor))

    # If that didn't work, try column-matching approach
    if not patients:
        # Extract all room numbers (3 digits, optionally followed by letter)
        rooms = re.findall(r'\b(\d{3})[A-Z]?\b', text)
        # Remove duplicates from room column headers like "343 343A"
        unique_rooms = []
        for r in rooms:
            if r not in unique_rooms:
                unique_rooms.append(r)

        # Extract all Med team numbers
        teams = re.findall(r'Med\s*(\d+)', text, re.IGNORECASE)

        # If counts match reasonably, zip them together
        if len(unique_rooms) > 0 and len(teams) > 0:
            # Repeat teams if there are more rooms than teams
            if len(unique_rooms) > len(teams):
                # Assume teams repeat in order
                extended_teams = (teams * ((len(unique_rooms) // len(teams)) + 1))[:len(unique_rooms)]
                teams = extended_teams

            for room, team in zip(unique_rooms, teams):
                if room not in seen_rooms:
                    seen_rooms.add(room)
                    floor = normalize_floor(room)
                    patients.append(ExistingPatient(room=room, current_team=int(team), floor=floor))

    return patients


def optimize_redistribution(
    existing_patients: list[ExistingPatient],
    closed_teams: set[int] = None
) -> list[tuple[ExistingPatient, int, str]]:
    """Optimize redistribution of existing patients to teams."""
    closed_teams = closed_teams or set()
    open_teams = [t for t in ALL_TEAMS if t not in closed_teams]
    regular_open_teams = [t for t in open_teams if t not in OVERFLOW_TEAMS]

    team_assignments = {t: [] for t in ALL_TEAMS}
    results = []

    # First pass: geographic assignments
    unassigned = []
    for patient in existing_patients:
        geo_teams = get_geographic_teams(patient.floor) if patient.floor else []
        geo_teams = [t for t in geo_teams if t not in closed_teams]

        if geo_teams:
            valid_teams = []
            for t in geo_teams:
                if t in IMCU_TEAMS:
                    if len(team_assignments[t]) < IMCU_CAP:
                        valid_teams.append(t)
                else:
                    valid_teams.append(t)

            if valid_teams:
                best_team = min(valid_teams, key=lambda t: len(team_assignments[t]))
                team_assignments[best_team].append(patient)
                reason = "Geographic" if best_team != patient.current_team else "No change"
                results.append((patient, best_team, reason))
                continue

        unassigned.append(patient)

    # Second pass: balance remaining
    for patient in unassigned:
        non_imcu = [t for t in regular_open_teams if t not in IMCU_TEAMS]
        if non_imcu:
            best_team = min(non_imcu, key=lambda t: len(team_assignments[t]))
            team_assignments[best_team].append(patient)
            results.append((patient, best_team, "Census balance"))

    return results


# =============================================================================
# STREAMLIT UI
# =============================================================================

st.set_page_config(
    page_title="Monday Shuffle - Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

st.title("Monday Shuffle")
st.markdown("Redistribute patients based on geography.")

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

# Input method tabs
input_tab1, input_tab2 = st.tabs(["Paste from Epic", "Upload Screenshots"])

all_patients = []

with input_tab1:
    st.markdown("**Enter room numbers only** - one per line. We'll assign teams based on geography.")
    st.markdown("*Just the room number: `304A`, `343B`, `534`, etc.*")

    paste_input = st.text_area(
        "Room numbers (one per line)",
        height=400,
        placeholder="304A\n343B\n534\n435\n..."
    )

    if paste_input and st.button("Assign Teams by Geography", type="primary", key="paste_btn"):
        seen_rooms = set()
        for line in paste_input.strip().split('\n'):
            line = line.strip()
            if not line:
                continue

            # Look for room number (3 digits with optional letter)
            room_match = re.search(r'\b(\d{3}[A-Z]?)\b', line, re.IGNORECASE)
            if not room_match:
                continue

            room = room_match.group(1).upper()

            if room in seen_rooms:
                continue
            seen_rooms.add(room)

            floor = normalize_floor(room)
            # Set current_team to 0 since we don't know it
            all_patients.append(ExistingPatient(room=room, current_team=0, floor=floor))

        st.success(f"Found {len(all_patients)} rooms")

with input_tab2:
    st.markdown("**Upload screenshots** of your Epic patient list.")

    if not OCR_AVAILABLE:
        st.warning("OCR not available on this server. Use the 'Paste from Epic' tab instead.")
    else:
        uploaded_files = st.file_uploader(
            "Upload Epic screenshots",
            type=['png', 'jpg', 'jpeg'],
            accept_multiple_files=True
        )

        if uploaded_files and st.button("Process Screenshots", type="primary", key="ocr_btn"):
            with st.spinner("Processing screenshots with OCR..."):
                for uploaded_file in uploaded_files:
                    image = Image.open(uploaded_file)

                    # Show raw OCR text for debugging
                    raw_text = pytesseract.image_to_string(image)
                    with st.expander(f"Raw OCR text from {uploaded_file.name}"):
                        st.code(raw_text)

                    patients = parse_epic_screenshot(image)
                    all_patients.extend(patients)
                    st.success(f"Found {len(patients)} patients in {uploaded_file.name}")

# Process results if we have patients
if all_patients:
    # Remove duplicates
    seen_rooms = set()
    unique_patients = []
    for p in all_patients:
        if p.room not in seen_rooms:
            seen_rooms.add(p.room)
            unique_patients.append(p)

    st.markdown(f"**Total unique patients: {len(unique_patients)}**")

    if unique_patients:
        results = optimize_redistribution(unique_patients, closed_teams)

        st.markdown("---")
        st.subheader("Redistribution Results")

        # Check if we have current team data
        has_current_teams = any(p.current_team != 0 for p, _, _ in results)

        if has_current_teams:
            changes = [(p, new_t, r) for p, new_t, r in results if new_t != p.current_team]
            no_changes = [(p, new_t, r) for p, new_t, r in results if new_t == p.current_team]

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Patients", len(results))
            col2.metric("Patients Moving", len(changes))
            col3.metric("Staying Put", len(no_changes))
        else:
            col1, col2 = st.columns(2)
            col1.metric("Total Patients", len(results))
            col2.metric("Teams Assigned", len(set(t for _, t, _ in results)))

        res_col1, res_col2 = st.columns(2)

        with res_col1:
            st.markdown("### Team Assignments")
            assign_text = ""
            for p, new_team, reason in sorted(results, key=lambda x: x[0].room):
                if p.current_team == 0:
                    assign_text += f"{p.room} → Med {new_team}\n"
                else:
                    marker = "→" if new_team != p.current_team else "="
                    assign_text += f"{p.room}: Med {p.current_team} {marker} Med {new_team}\n"
            st.code(assign_text, language=None)

        with res_col2:
            st.markdown("### New Census by Team")
            team_counts = defaultdict(int)
            for p, new_team, reason in results:
                team_counts[new_team] += 1

            census_text = ""
            for team in ALL_TEAMS:
                if team in closed_teams:
                    census_text += f"Med {team:2d}: CLOSED\n"
                else:
                    count = team_counts.get(team, 0)
                    imcu = "*" if team in IMCU_TEAMS else " "
                    census_text += f"Med {team:2d}{imcu}: {count:2d} patients\n"
            census_text += "\n* = IMCU"
            st.code(census_text, language=None)

        st.markdown("### Full Assignment List")
        full_text = "Room     Current  →  New      Reason\n"
        full_text += "-" * 45 + "\n"
        for p, new_team, reason in sorted(results, key=lambda x: x[0].room):
            change_marker = "→" if new_team != p.current_team else "="
            full_text += f"{p.room:8} Med {p.current_team:2d}  {change_marker}  Med {new_team:2d}   {reason}\n"
        st.code(full_text, language=None)
