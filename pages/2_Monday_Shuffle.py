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
    """
    Analyze patients and separate into wrong team vs OK.

    Returns:
        wrong_team: List of (patient, acceptable_teams) for patients needing reassignment
        ok_team: List of patients already on acceptable teams
    """
    closed_teams = closed_teams or set()

    wrong_team = []
    ok_team = []

    for patient in patients:
        geo_teams = get_geographic_teams(patient.floor) if patient.floor else []
        # Filter out closed teams from acceptable options
        acceptable_teams = [t for t in geo_teams if t not in closed_teams]

        if patient.current_team in acceptable_teams:
            # Patient is already on an acceptable geographic team
            ok_team.append(patient)
        elif acceptable_teams:
            # Patient is on wrong team, has geographic options
            wrong_team.append((patient, acceptable_teams))
        else:
            # No geographic teams (unknown floor or all geo teams closed)
            # Consider them "wrong" with empty acceptable list
            wrong_team.append((patient, []))

    return wrong_team, ok_team


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

st.markdown("""
**Enter room and current team** - one per line.

Format: `Room Team` (e.g., `304A 1` or `304A Med 1`)

Examples:
```
304A 1
343B 5
534 Med 10
435A 7
```
""")

paste_input = st.text_area(
    "Room and Team (one per line)",
    height=400,
    placeholder="304A 1\n343B 5\n534 Med 10\n435A 7\n..."
)

if st.button("Analyze Teams", type="primary"):
    if not paste_input:
        st.warning("Please enter room and team data first.")
    else:
        all_patients = []
        seen_rooms = set()
        parse_errors = []

        for line in paste_input.strip().split('\n'):
            line = line.strip()
            if not line:
                continue

            # Look for room number (3 digits with optional letter)
            room_match = re.search(r'\b(\d{3}[A-Z]?)\b', line, re.IGNORECASE)
            if not room_match:
                parse_errors.append(f"No room found: {line}")
                continue

            room = room_match.group(1).upper()

            # Look for team number (with or without "Med" prefix)
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
            all_patients.append(ExistingPatient(room=room, current_team=team_num, floor=floor))

        if parse_errors:
            with st.expander(f"Parse warnings ({len(parse_errors)})"):
                for err in parse_errors:
                    st.warning(err)

        st.success(f"Parsed {len(all_patients)} patients")

        if all_patients:
            wrong_team, ok_team = analyze_patients(all_patients, closed_teams)

            st.markdown("---")
            st.subheader("Analysis Results")

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Patients", len(all_patients))
            col2.metric("Need Reassignment", len(wrong_team))
            col3.metric("Already OK", len(ok_team))

            res_col1, res_col2 = st.columns(2)

            with res_col1:
                st.markdown("### Needs Reassignment")
                if wrong_team:
                    wrong_text = ""
                    for patient, acceptable in sorted(wrong_team, key=lambda x: x[0].room):
                        if acceptable:
                            options = ", ".join(f"Med {t}" for t in acceptable)
                            wrong_text += f"{patient.room}: Med {patient.current_team} -> {options}\n"
                        else:
                            wrong_text += f"{patient.room}: Med {patient.current_team} -> ? (no geo match)\n"
                    st.code(wrong_text, language=None)
                else:
                    st.info("All patients are on acceptable teams!")

            with res_col2:
                st.markdown("### Already on Correct Team")
                if ok_team:
                    ok_text = ""
                    for patient in sorted(ok_team, key=lambda x: x.room):
                        ok_text += f"{patient.room}: Med {patient.current_team} OK\n"
                    st.code(ok_text, language=None)
                else:
                    st.info("No patients on correct teams yet.")

            # Summary by current team
            st.markdown("### Current Census by Team")
            team_counts = defaultdict(list)
            for patient in all_patients:
                team_counts[patient.current_team].append(patient.room)

            census_text = ""
            for team in ALL_TEAMS:
                if team in closed_teams:
                    census_text += f"Med {team:2d}: CLOSED\n"
                else:
                    rooms = team_counts.get(team, [])
                    census_text += f"Med {team:2d}: {len(rooms):2d} patients\n"
            st.code(census_text, language=None)

            # Detailed wrong team list for copying
            if wrong_team:
                st.markdown("### Reassignment List (for Epic)")
                epic_text = "Room       From    To Options\n"
                epic_text += "-" * 40 + "\n"
                for patient, acceptable in sorted(wrong_team, key=lambda x: x[0].room):
                    if acceptable:
                        options = "/".join(str(t) for t in acceptable)
                        epic_text += f"{patient.room:10} Med {patient.current_team:2d}  Med {options}\n"
                    else:
                        epic_text += f"{patient.room:10} Med {patient.current_team:2d}  ? (check location)\n"
                st.code(epic_text, language=None)
