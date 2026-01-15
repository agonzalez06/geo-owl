#!/usr/bin/env python3
"""
Geographic Patient Placement Optimizer - Web Version (Streamlit)

Run locally: streamlit run geo_placer_web.py
"""

import streamlit as st
import streamlit.components.v1 as components
import re
from dataclasses import dataclass
from typing import Optional
from collections import defaultdict
from PIL import Image
import pytesseract
import io

# =============================================================================
# GEOGRAPHIC MAPPINGS (same as CLI version)
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

TEAM_FLOORS = {
    1: ['3W', '3E', 'IMCU'], 2: ['3W', '3E', 'IMCU'], 3: ['3W', '3E', 'IMCU'],
    4: ['4E', '4W'], 5: ['5E', '5W'], 6: ['6E', '6W'],
    7: ['7E', '7W'], 8: ['8E'], 9: ['7E', '7W'],
    10: ['5E', '5W'], 11: ['4E', '4W'], 12: ['6E', '6W', 'Boyer'],
    13: ['8E'], 14: ['Overflow'], 15: ['Overflow'],
}

ALL_TEAMS = list(range(1, 16))
OVERFLOW_TEAMS = [14, 15]
IMCU_TEAMS = [1, 2, 3]
IMCU_CAP = 10
SOFT_CAP = 14
MAX_NEW_BEFORE_SPREAD = 3
MAX_CENSUS_GAP = 4


@dataclass
class Patient:
    identifier: str
    floor: str
    raw_location: str
    admitted_by: str = ""  # Overnight doctor who admitted


@dataclass
class Assignment:
    patient: Patient
    team: int
    is_geographic: bool
    reason: str


@dataclass
class ExistingPatient:
    """Patient with current team assignment (for Monday redistribution)."""
    room: str
    current_team: int
    floor: str


def parse_epic_screenshot(image) -> list[ExistingPatient]:
    """
    Parse Epic patient list screenshot using OCR.
    Returns list of patients with room and current team.
    """
    # Run OCR
    text = pytesseract.image_to_string(image)

    patients = []
    seen_rooms = set()

    # Pattern to match room and Med team
    # Examples: "3W 304A Med 1", "5E 534A Med 5", "4W 410A Med 4"
    # Room pattern: floor+wing + room number (e.g., "3W 304A" or "304A")

    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue

        # Look for Med team number
        team_match = re.search(r'Med\s*(\d+)', line, re.IGNORECASE)
        if not team_match:
            continue

        team_num = int(team_match.group(1))

        # Look for room number (3-digit with optional letter suffix)
        room_match = re.search(r'\b(\d{3}[A-Z]?)\b', line)
        if not room_match:
            continue

        room = room_match.group(1)

        # Skip duplicates
        if room in seen_rooms:
            continue
        seen_rooms.add(room)

        # Normalize floor from room number
        floor = normalize_floor(room)

        patients.append(ExistingPatient(
            room=room,
            current_team=team_num,
            floor=floor
        ))

    return patients


def optimize_redistribution(
    existing_patients: list[ExistingPatient],
    closed_teams: set[int] = None
) -> list[tuple[ExistingPatient, int, str]]:
    """
    Optimize redistribution of existing patients to teams.
    Returns list of (patient, new_team, reason) tuples.
    Only suggests changes for patients who should move.
    """
    closed_teams = closed_teams or set()
    open_teams = [t for t in ALL_TEAMS if t not in closed_teams]
    regular_open_teams = [t for t in open_teams if t not in OVERFLOW_TEAMS]

    # Calculate target census per team (even distribution)
    total_patients = len(existing_patients)
    num_teams = len(regular_open_teams)
    target_per_team = total_patients // num_teams if num_teams > 0 else 0

    # Track assignments
    team_assignments = {t: [] for t in ALL_TEAMS}
    results = []

    # First pass: assign patients to geographic teams where possible
    unassigned = []
    for patient in existing_patients:
        geo_teams = get_geographic_teams(patient.floor) if patient.floor else []
        geo_teams = [t for t in geo_teams if t not in closed_teams]

        if geo_teams:
            # Check IMCU caps
            valid_teams = []
            for t in geo_teams:
                if t in IMCU_TEAMS:
                    current_count = len(team_assignments[t])
                    if current_count < IMCU_CAP:
                        valid_teams.append(t)
                else:
                    valid_teams.append(t)

            if valid_teams:
                # Pick team with lowest current assignment count
                best_team = min(valid_teams, key=lambda t: len(team_assignments[t]))
                team_assignments[best_team].append(patient)

                reason = "Geographic" if best_team != patient.current_team else "No change (geographic)"
                results.append((patient, best_team, reason))
                continue

        unassigned.append(patient)

    # Second pass: assign remaining patients to balance census
    for patient in unassigned:
        non_imcu = [t for t in regular_open_teams if t not in IMCU_TEAMS]
        if non_imcu:
            best_team = min(non_imcu, key=lambda t: len(team_assignments[t]))
            team_assignments[best_team].append(patient)
            results.append((patient, best_team, "Census balance"))

    return results


# =============================================================================
# FLOOR PARSING (same as CLI version)
# =============================================================================

def normalize_floor(location: str) -> Optional[str]:
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

    match = re.search(r'FLOOR\s*(\d+)\s*(EAST|WEST|E|W)', location)
    if match:
        floor_num = match.group(1)
        direction = match.group(2)[0]
        return f"{floor_num}{direction}"

    return None


def get_geographic_teams(floor: str) -> list[int]:
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


# =============================================================================
# PLACEMENT ALGORITHM (same as CLI version)
# =============================================================================

def optimize_placements(
    patients: list[Patient],
    current_census: dict[int, int],
    closed_teams: set[int] = None
) -> list[Assignment]:
    assignments = []
    census = current_census.copy()
    new_assignments = {t: 0 for t in ALL_TEAMS}
    closed_teams = closed_teams or set()

    open_teams = [t for t in ALL_TEAMS if t not in closed_teams]
    regular_open_teams = [t for t in open_teams if t not in OVERFLOW_TEAMS]

    patients_sorted = sorted(patients, key=lambda p: (p.floor or 'ZZZ', p.identifier))

    for patient in patients_sorted:
        geo_teams_raw = get_geographic_teams(patient.floor) if patient.floor else []
        geo_teams = [t for t in geo_teams_raw if t not in closed_teams]
        best_team = None
        is_geo = False
        reason = ""

        if geo_teams:
            valid_geo_teams = []
            for t in geo_teams:
                if t in IMCU_TEAMS and census.get(t, 0) >= IMCU_CAP:
                    continue
                if t not in IMCU_TEAMS and census.get(t, 0) >= SOFT_CAP:
                    others_under_cap = any(
                        census.get(other, 0) < SOFT_CAP
                        and (other not in IMCU_TEAMS or census.get(other, 0) < IMCU_CAP)
                        for other in geo_teams if other != t
                    )
                    if others_under_cap:
                        continue
                if new_assignments[t] >= MAX_NEW_BEFORE_SPREAD:
                    others_have_less = any(
                        new_assignments[other] < new_assignments[t]
                        and (other not in IMCU_TEAMS or census.get(other, 0) < IMCU_CAP)
                        and census.get(other, 0) < SOFT_CAP
                        for other in geo_teams if other != t
                    )
                    if others_have_less:
                        continue
                valid_geo_teams.append(t)

            if valid_geo_teams:
                best_geo_team = min(valid_geo_teams, key=lambda t: census.get(t, 0))
                best_geo_census = census.get(best_geo_team, 0)

                non_imcu_regular = [t for t in regular_open_teams if t not in IMCU_TEAMS]
                if non_imcu_regular:
                    lowest_census_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
                    lowest_census = census.get(lowest_census_team, 0)

                    if best_geo_census >= lowest_census + MAX_CENSUS_GAP and lowest_census_team not in valid_geo_teams:
                        best_team = lowest_census_team
                        is_geo = False
                        reason = f"Balance override ({patient.floor}→Med {best_geo_team} would be {best_geo_census+1}, Med {lowest_census_team} only {lowest_census})"
                    else:
                        best_team = best_geo_team
                        is_geo = True
                        reason = f"Geographic ({patient.floor} → Med {best_team})"
                else:
                    best_team = best_geo_team
                    is_geo = True
                    reason = f"Geographic ({patient.floor} → Med {best_team})"
            elif geo_teams:
                available_geo = [t for t in geo_teams if t not in IMCU_TEAMS or census.get(t, 0) < IMCU_CAP]

                if available_geo:
                    best_geo = min(available_geo, key=lambda t: census.get(t, 0))
                    best_geo_census = census.get(best_geo, 0)

                    non_imcu_non_geo = [t for t in regular_open_teams if t not in IMCU_TEAMS and t not in geo_teams]
                    if non_imcu_non_geo:
                        best_non_geo = min(non_imcu_non_geo, key=lambda t: census.get(t, 0))
                        best_non_geo_census = census.get(best_non_geo, 0)

                        if best_geo_census >= best_non_geo_census + 3:
                            best_team = best_non_geo
                            is_geo = False
                            reason = f"Balance override ({patient.floor} geo full, Med {best_team} lower census)"
                        else:
                            best_team = best_geo
                            is_geo = True
                            reason = f"Geographic (over soft cap, {patient.floor} → Med {best_team})"
                    else:
                        best_team = best_geo
                        is_geo = True
                        reason = f"Geographic (equity override, {patient.floor} → Med {best_team})"

        if best_team is None:
            non_imcu_regular = [t for t in regular_open_teams if t not in IMCU_TEAMS]
            non_imcu_overflow = [t for t in OVERFLOW_TEAMS if t in open_teams and t not in IMCU_TEAMS]

            regular_under_cap = [t for t in non_imcu_regular if census.get(t, 0) < SOFT_CAP]
            if regular_under_cap:
                best_team = min(regular_under_cap, key=lambda t: census.get(t, 0))
            elif non_imcu_overflow:
                overflow_under_cap = [t for t in non_imcu_overflow if census.get(t, 0) < SOFT_CAP]
                if overflow_under_cap:
                    best_team = min(overflow_under_cap, key=lambda t: census.get(t, 0))
                elif non_imcu_regular:
                    best_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
                else:
                    best_team = min(non_imcu_overflow, key=lambda t: census.get(t, 0))
            elif non_imcu_regular:
                best_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
            elif open_teams:
                best_team = min(open_teams, key=lambda t: census.get(t, 0))
            else:
                continue

            is_geo = False
            if patient.floor == 'BOYER':
                reason = f"Boyer overflow (Med 12 full), lowest census"
            elif patient.floor:
                reason = f"No geographic capacity for {patient.floor}, lowest census"
            else:
                reason = f"No floor specified, lowest census"

        census[best_team] = census.get(best_team, 0) + 1
        new_assignments[best_team] += 1

        assignments.append(Assignment(
            patient=patient,
            team=best_team,
            is_geographic=is_geo,
            reason=reason
        ))

    return assignments


# =============================================================================
# STREAMLIT UI
# =============================================================================

st.set_page_config(
    page_title="Geo Owl",
    page_icon="Gemini_Generated_Image_2hkaog2hkaog2hka.png",
    layout="wide"
)

# JavaScript to make Enter key move to next input field
components.html("""
<script>
const doc = window.parent.document;
doc.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
        e.preventDefault();
        const inputs = Array.from(doc.querySelectorAll('input[type="text"]'));
        const currentIndex = inputs.indexOf(e.target);
        if (currentIndex < inputs.length - 1) {
            inputs[currentIndex + 1].focus();
        }
    }
});
</script>
""", height=0)

# Temple University red for buttons
st.markdown("""
<style>
    .stButton > button[kind="primary"] {
        background-color: #9D2235;
        border-color: #9D2235;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #7A1A2A;
        border-color: #7A1A2A;
    }
</style>
""", unsafe_allow_html=True)

# Header with logo
header_col1, header_col2 = st.columns([1, 8])
with header_col1:
    st.image("Gemini_Generated_Image_2hkaog2hkaog2hka.png", use_container_width=True)
with header_col2:
    st.title("Geographic Placement Optimizer")
    st.markdown("Optimal team assignments based on geography and census.")

# Mode selector
mode = st.radio(
    "Mode",
    ["T-List (Overnight Admissions)", "Monday Shuffle (Redistribution)"],
    horizontal=True,
    label_visibility="collapsed"
)

if mode == "T-List (Overnight Admissions)":
    # =============================================================================
    # T-LIST MODE
    # =============================================================================

    # Create 5 columns: Census + 4 Doctors side by side
    census_col, doc1_col, doc2_col, doc3_col, doc4_col = st.columns([1, 1, 1, 1, 1])

# Column 1: Team Census (compact layout)
with census_col:
    st.subheader("Census")
    st.caption("X = closed")

    census = {}
    closed_teams = set()

    for team in ALL_TEAMS:
        label_col, input_col = st.columns([1, 3])
        with label_col:
            imcu = "*" if team in IMCU_TEAMS else ""
            st.markdown(f"<div style='padding-top:8px; text-align:right'>Med {team}{imcu}</div>", unsafe_allow_html=True)
        with input_col:
            value = st.text_input(
                f"Med {team}",
                key=f"census_{team}",
                label_visibility="collapsed"
            )

        if value:
            value_upper = value.strip().upper()
            if value_upper in ('X', 'NA', 'CLOSED', 'N/A', '-'):
                closed_teams.add(team)
                census[team] = 0
            else:
                try:
                    census[team] = int(value)
                except ValueError:
                    census[team] = 0
        else:
            census[team] = 0

# Columns 2-5: Overnight Doctors (Amion naming)
doc_cols = [doc1_col, doc2_col, doc3_col, doc4_col]
doc_labels = [
    ("Med Q", "1-3"),
    ("Med S", "4-6"),
    ("Med Y", "7-9"),
    ("Med Z", "10-13"),
]
doctor_names = []
doctor_patients = []

for i, (doc_col, (code, teams)) in enumerate(zip(doc_cols, doc_labels), 1):
    with doc_col:
        st.subheader(f"{code} ({teams})")
        name = st.text_input(
            "Name",
            key=f"doc_{i}",
            placeholder="Name",
            label_visibility="collapsed"
        )
        doctor_names.append(name.strip() if name else code)

        patients = st.text_area(
            "Patients",
            key=f"patients_{i}",
            height=400,
            placeholder="312\n545\n7E\n...",
            label_visibility="collapsed"
        )
        doctor_patients.append(patients)

# Show closed teams
if closed_teams:
    st.info(f"**Closed teams:** {', '.join(f'Med {t}' for t in sorted(closed_teams))}")

# Process button
if st.button("Optimize Placements", type="primary", use_container_width=True):
    # Parse patients from all 4 doctor columns
    patients = []
    seen = set()
    duplicates = 0
    pt_count = 1

    for doc_idx, (doc_name, doc_patients_text) in enumerate(zip(doctor_names, doctor_patients)):
        if not doc_patients_text:
            continue

        for line in doc_patients_text.strip().split('\n'):
            location = line.strip()
            if not location:
                continue

            location_key = location.upper()
            if location_key in seen:
                duplicates += 1
                continue
            seen.add(location_key)

            floor = normalize_floor(location)
            patients.append(Patient(f"Pt{pt_count}", floor, location, admitted_by=doc_name))
            pt_count += 1

    if not patients:
        st.warning("No patients entered. Please enter patient locations above.")
    else:
        # Run optimization
        assignments = optimize_placements(patients, census, closed_teams)

        # Calculate final census
        final_census = census.copy()
        for a in assignments:
            final_census[a.team] = final_census.get(a.team, 0) + 1

        # Show results with anchor for scrolling
        st.markdown("---")
        st.markdown("<div id='results-section'></div>", unsafe_allow_html=True)
        st.subheader("Results")

        # Auto-scroll to results
        components.html("""
            <script>
                window.parent.document.getElementById('results-section').scrollIntoView({behavior: 'smooth'});
            </script>
        """, height=0)

        if duplicates > 0:
            st.warning(f"Skipped {duplicates} duplicate(s)")

        geo_count = sum(1 for a in assignments if a.is_geographic)
        total = len(assignments)

        # Summary metrics
        metric_cols = st.columns(4)
        metric_cols[0].metric("Total Patients", total)
        metric_cols[1].metric("Geographic", f"{geo_count} ({100*geo_count//total if total else 0}%)")
        metric_cols[2].metric("Non-Geographic", total - geo_count)
        metric_cols[3].metric("Teams Used", len(set(a.team for a in assignments)))

        # Results in 4 columns
        res_col1, res_col2, res_col3, res_col4 = st.columns(4)

        # Column 1: Census Summary
        with res_col1:
            st.markdown("### Census Summary")

            summary_text = "Team   Start +New =Final\n"
            summary_text += "-" * 26 + "\n"

            for team in ALL_TEAMS:
                if team in closed_teams:
                    summary_text += f"Med {team:2d}   -- CLOSED\n"
                    continue

                start = census.get(team, 0)
                final = final_census.get(team, 0)
                new = final - start

                warning = ""
                if team in IMCU_TEAMS and final >= IMCU_CAP:
                    warning = " CAP"
                elif team not in IMCU_TEAMS and final >= SOFT_CAP:
                    warning = " HIGH"

                imcu = "*" if team in IMCU_TEAMS else " "
                summary_text += f"Med {team:2d}{imcu} {start:2d}  +{new:2d}  ={final:2d}{warning}\n"

            summary_text += "\n* = IMCU (cap: 10)"
            st.code(summary_text, language=None)

        # Column 2: Assignment List
        with res_col2:
            st.markdown("### Assignment List")

            # Sort by room number/location
            sorted_assignments = sorted(assignments, key=lambda a: a.patient.raw_location)
            assignment_text = ""
            for a in sorted_assignments:
                assignment_text += f"{a.patient.raw_location:8} → Med {a.team:2d} ({a.patient.admitted_by})\n"

            st.code(assignment_text, language=None)

        # Column 3: By Team (monospace)
        with res_col3:
            st.markdown("### By Team")

            by_team = defaultdict(list)
            for a in assignments:
                by_team[a.team].append(a)

            by_team_text = ""
            for team in ALL_TEAMS:
                if team in closed_teams:
                    continue

                team_assignments = by_team[team]
                if not team_assignments:
                    continue

                start = census.get(team, 0)
                final = final_census.get(team, 0)
                new_count = len(team_assignments)

                imcu = "*" if team in IMCU_TEAMS else ""
                by_team_text += f"Med {team}{imcu} ({start}→{final})\n"

                for a in team_assignments:
                    by_team_text += f"  {a.patient.raw_location} ({a.patient.admitted_by})\n"

                by_team_text += "\n"

            st.code(by_team_text, language=None)

        # Column 4: EPIC Secure Chat Message
        with res_col4:
            st.markdown("### EPIC Message")

            by_team = defaultdict(list)
            for a in assignments:
                by_team[a.team].append((a.patient.raw_location, a.patient.admitted_by))

            epic_lines = ["Good Morning! Here are today's redis:"]
            for team in sorted(by_team.keys()):
                patient_strs = [f"{room} ({doc})" for room, doc in by_team[team]]
                epic_lines.append(f"Med {team}: {', '.join(patient_strs)}")

            epic_message = "\n".join(epic_lines)

            # Calculate height based on lines (approx 20px per line + padding)
            line_count = len(epic_lines)
            text_height = max(150, line_count * 22 + 50)

            # Display with copy button using HTML/JS
            escaped_message = epic_message.replace('`', '\\`').replace('$', '\\$')
            components.html(f"""
                <style>
                    .epic-container {{
                        font-family: monospace;
                        background-color: #f0f2f6;
                        border-radius: 5px;
                        padding: 10px;
                        white-space: pre-wrap;
                        font-size: 14px;
                        line-height: 1.4;
                    }}
                    .copy-btn {{
                        background-color: #9D2235;
                        color: white;
                        border: none;
                        padding: 8px 16px;
                        border-radius: 5px;
                        cursor: pointer;
                        margin-bottom: 10px;
                        font-weight: bold;
                    }}
                    .copy-btn:hover {{
                        background-color: #7A1A2A;
                    }}
                </style>
                <button class="copy-btn" onclick="copyToClipboard()">Copy Message</button>
                <div class="epic-container" id="epicMsg">{epic_message}</div>
                <script>
                    function copyToClipboard() {{
                        const text = document.getElementById('epicMsg').innerText;
                        navigator.clipboard.writeText(text).then(() => {{
                            const btn = document.querySelector('.copy-btn');
                            btn.innerText = 'Copied!';
                            setTimeout(() => {{ btn.innerText = 'Copy Message'; }}, 2000);
                        }});
                    }}
                </script>
            """, height=text_height + 60)

else:
    # =============================================================================
    # MONDAY SHUFFLE MODE
    # =============================================================================

    st.markdown("### Upload Epic Screenshots")
    st.markdown("Upload screenshots of your Epic patient list. The OCR will extract room numbers and current team assignments.")

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

    # File uploader for multiple images
    uploaded_files = st.file_uploader(
        "Upload Epic screenshots",
        type=['png', 'jpg', 'jpeg'],
        accept_multiple_files=True
    )

    if uploaded_files and st.button("Process Screenshots & Optimize", type="primary"):
        all_patients = []

        with st.spinner("Processing screenshots with OCR..."):
            for uploaded_file in uploaded_files:
                image = Image.open(uploaded_file)
                patients = parse_epic_screenshot(image)
                all_patients.extend(patients)
                st.success(f"Found {len(patients)} patients in {uploaded_file.name}")

        # Remove duplicates
        seen_rooms = set()
        unique_patients = []
        for p in all_patients:
            if p.room not in seen_rooms:
                seen_rooms.add(p.room)
                unique_patients.append(p)

        st.markdown(f"**Total unique patients: {len(unique_patients)}**")

        if unique_patients:
            # Run optimization
            results = optimize_redistribution(unique_patients, closed_teams)

            # Show results
            st.markdown("---")
            st.subheader("Redistribution Results")

            # Count changes
            changes = [(p, new_t, r) for p, new_t, r in results if new_t != p.current_team]
            no_changes = [(p, new_t, r) for p, new_t, r in results if new_t == p.current_team]

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Patients", len(results))
            col2.metric("Patients Moving", len(changes))
            col3.metric("Staying Put", len(no_changes))

            # Results in columns
            res_col1, res_col2 = st.columns(2)

            with res_col1:
                st.markdown("### Patients to Move")
                if changes:
                    move_text = ""
                    for p, new_team, reason in sorted(changes, key=lambda x: x[0].room):
                        move_text += f"{p.room}: Med {p.current_team} → Med {new_team}\n"
                    st.code(move_text, language=None)
                else:
                    st.info("No patients need to move!")

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

            # Full assignment list
            st.markdown("### Full Assignment List")
            full_text = "Room     Current  →  New      Reason\n"
            full_text += "-" * 45 + "\n"
            for p, new_team, reason in sorted(results, key=lambda x: x[0].room):
                change_marker = "→" if new_team != p.current_team else "="
                full_text += f"{p.room:8} Med {p.current_team:2d}  {change_marker}  Med {new_team:2d}   {reason}\n"
            st.code(full_text, language=None)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "Room convention: X01-X20 = West, X30-X50 = East | "
    "IMCU cap: 10 | Soft cap: 14"
    "</div>",
    unsafe_allow_html=True
)
