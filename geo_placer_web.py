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
from PIL import Image, ImageFilter, ImageEnhance, ImageOps

try:
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


def preprocess_image_for_ocr(image):
    """
    Preprocess image to improve OCR accuracy, especially for camera photos.
    Handles rotation, moiré patterns, and contrast issues.
    """
    # Try to auto-detect and fix orientation using EXIF data
    try:
        image = ImageOps.exif_transpose(image)
    except Exception:
        pass

    # Convert to grayscale
    gray = image.convert('L')

    # For camera photos: apply slight blur to reduce moiré patterns
    blurred = gray.filter(ImageFilter.GaussianBlur(radius=0.5))

    # Enhance contrast
    enhancer = ImageEnhance.Contrast(blurred)
    enhanced = enhancer.enhance(1.5)

    # Try to detect if image needs rotation using Tesseract OSD
    try:
        osd = pytesseract.image_to_osd(enhanced)
        rotation_match = re.search(r'Rotate: (\d+)', osd)
        if rotation_match:
            rotation = int(rotation_match.group(1))
            if rotation != 0:
                # Tesseract reports how much to rotate to fix, so we rotate
                enhanced = enhanced.rotate(-rotation, expand=True)
    except Exception:
        # OSD can fail on some images, try common rotations manually
        pass

    return enhanced


def try_all_rotations(image):
    """
    Try OCR on all 4 rotations and return the best result.
    Useful when OSD detection fails.
    """
    best_pairs = []
    best_text = ""
    best_rotation = 0
    best_psm = 6
    last_error = None
    any_success = False

    # Resize large images for faster OCR (max 1500px on longest side)
    max_dimension = 1500
    if max(image.size) > max_dimension:
        ratio = max_dimension / max(image.size)
        new_size = (int(image.size[0] * ratio), int(image.size[1] * ratio))
        image = image.resize(new_size, Image.LANCZOS)

    gray = image.convert('L')

    # Apply preprocessing
    blurred = gray.filter(ImageFilter.GaussianBlur(radius=0.5))
    enhancer = ImageEnhance.Contrast(blurred)
    enhanced = enhancer.enhance(1.5)

    # Try 0° and 90° (covers landscape/portrait), with 2 best PSM modes
    for rotation in [0, 90]:
        if rotation == 0:
            rotated = enhanced
        else:
            rotated = enhanced.rotate(-rotation, expand=True)

        for psm in [6, 4]:
            config = f'--psm {psm}'
            try:
                raw_text = pytesseract.image_to_string(rotated, config=config)
                any_success = True
                pairs = extract_from_ocr(raw_text)
                if len(pairs) > len(best_pairs):
                    best_pairs = pairs
                    best_text = raw_text
                    best_rotation = rotation
                    best_psm = psm
                # Keep best raw text even if no pairs found
                if not best_text and raw_text:
                    best_text = raw_text
            except Exception as e:
                last_error = str(e)
                continue

    # If no success at all, raise the last error
    if not any_success and last_error:
        raise Exception(f"OCR failed: {last_error}")

    return best_pairs, best_text, best_rotation, best_psm

# =============================================================================
# GEOGRAPHIC MAPPINGS
# =============================================================================

FLOOR_TO_TEAMS = {
    '3W': [1, 2, 3], '3E': [1, 2, 3],
    '4W': [4, 11], '4E': [4, 11],
    '5W': [5, 10], '5E': [5, 10],
    '6W': [6, 12], '6E': [6, 12],
    '7W': [7, 9], '7E': [7, 9],
    '8E': [8], '8W': [8],
    'IMCU': [1, 2, 3],
    'BOYER': [12, 6],
}

TEAM_FLOORS = {
    1: ['3W', '3E', 'IMCU'], 2: ['3W', '3E', 'IMCU'], 3: ['3W', '3E', 'IMCU'],
    4: ['4E', '4W'], 5: ['5E', '5W'], 6: ['6E', '6W'],
    7: ['7E', '7W'], 8: ['8E', '8W'], 9: ['7E', '7W'],
    10: ['5E', '5W'], 11: ['4E', '4W'], 12: ['6E', '6W', 'Boyer'],
    13: ['Overflow'], 14: ['Overflow'], 15: ['Overflow'],
}

TEAM_FLOORS_STR = {
    1: "3E/3W/IMCU", 2: "3E/3W/IMCU", 3: "3E/3W/IMCU",
    4: "4E/4W", 5: "5E/5W", 6: "6E/6W/Boyer",
    7: "7E/7W", 8: "8E/8W", 9: "7E/7W",
    10: "5E/5W", 11: "4E/4W", 12: "6E/6W/Boyer",
    13: "Overflow", 14: "Overflow", 15: "Overflow",
}

ALL_TEAMS = list(range(1, 16))
OVERFLOW_TEAMS = [13, 14, 15]
IMCU_TEAMS = [1, 2, 3]
IMCU_TARGET = 9
IMCU_CAP = 10
SOFT_CAP = 14
MAX_NEW_BEFORE_SPREAD = 3
MAX_CENSUS_GAP = 4


@dataclass
class Patient:
    identifier: str
    floor: str
    raw_location: str
    admitted_by: str = ""


@dataclass
class Assignment:
    patient: Patient
    team: int
    is_geographic: bool
    reason: str


@dataclass
class ExistingPatient:
    room: str
    current_team: int
    floor: str


# =============================================================================
# FLOOR PARSING
# =============================================================================

def normalize_floor(location: str) -> Optional[str]:
    original = location.strip().upper()

    if 'IMCU' in original:
        return 'IMCU'
    if 'OVERNIGHT' in original or 'ONR' in original or 'RECOVERY' in original:
        return 'BOYER'
    if original.startswith('RZ'):
        return 'BOYER'
    if original.startswith('Y'):
        return 'BOYER'
    if any(x in original for x in ['ED', 'EMERGENCY', 'ER ']):
        return None
    if 'BOYER' in original:
        return 'BOYER'
    if original == 'MAIN':
        return None

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
# OVERNIGHT PLACEMENT ALGORITHM
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
                def geo_score(t):
                    c = census.get(t, 0)
                    if t in IMCU_TEAMS and c >= IMCU_TARGET:
                        return c + 50
                    return c

                best_geo_team = min(valid_geo_teams, key=geo_score)
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
# MONDAY SHUFFLE FUNCTIONS
# =============================================================================

def analyze_patients(
    patients: list[ExistingPatient],
    closed_teams: set[int] = None
) -> tuple[list[tuple[ExistingPatient, list[int]]], list[ExistingPatient]]:
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
    pairs = []
    seen_rooms = set()
    room_pattern = r'\b(\d{3}[A-Z]?|[A-Z]{2,4}\d{0,2}[A-Z]?|\d{4}[A-Z]?)\b'

    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if 'Primary' in line or 'Bed' in line and 'Team' in line:
            continue

        room_match = re.search(room_pattern, line, re.IGNORECASE)
        team_match = re.search(r'Med\s*(\d{1,2})', line, re.IGNORECASE)

        if room_match and team_match:
            room = room_match.group(1).upper()
            team = int(team_match.group(1))
            if room in ('BED', 'PRIMARY', 'TEAM'):
                continue
            if room not in seen_rooms and 1 <= team <= 15:
                seen_rooms.add(room)
                pairs.append((room, team))

    all_rooms = re.findall(r'\b(\d{3,4}[A-Z]?)\b', text, re.IGNORECASE)
    special_rooms = re.findall(r'\b([A-Z]{2,4}\d{1,2})\b', text)
    special_rooms += re.findall(r'\b(MAIN|ICU\d*|IMCU\d*)\b', text, re.IGNORECASE)

    unique_rooms = []
    for r in all_rooms + special_rooms:
        r_upper = r.upper()
        if r_upper not in unique_rooms and r_upper not in ('BED', 'PRIMARY', 'TEAM', 'MED'):
            unique_rooms.append(r_upper)

    all_teams = re.findall(r'Med\s*(\d{1,2})', text, re.IGNORECASE)
    all_teams = [int(t) for t in all_teams if 1 <= int(t) <= 15]

    if len(pairs) >= 5 and abs(len(pairs) - len(all_teams)) <= 3:
        return pairs

    if len(unique_rooms) > 0 and len(all_teams) > 0:
        min_len = min(len(unique_rooms), len(all_teams))
        for room, team in zip(unique_rooms[:min_len], all_teams[:min_len]):
            if room not in seen_rooms:
                seen_rooms.add(room)
                pairs.append((room, team))

    return pairs


def fix_ocr_room(room: str) -> str:
    room = room.upper()
    garbage = ['MED', 'BED', 'TEAM', 'PRIMARY', 'POSE', 'TAA', 'ATTA']
    if room in garbage:
        return None

    if len(room) == 4 and room.isdigit():
        last = room[-1]
        if last == '4':
            room = room[:-1] + 'A'
        elif last == '8':
            room = room[:-1] + 'B'

    if room.startswith('T') and len(room) >= 3:
        fixed = '7' + room[1:]
        if re.match(r'^\d{3}[A-Z]?$', fixed):
            room = fixed

    return room


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

# Geography reference (collapsible)
with st.expander("Team Geography Reference"):
    ref_text = """
Team    Floors              Team    Floors
----    ------              ----    ------
Med 1*  3E / 3W / IMCU      Med 9   7E / 7W
Med 2*  3E / 3W / IMCU      Med 10  5E / 5W
Med 3*  3E / 3W / IMCU      Med 11  4E / 4W
Med 4   4E / 4W             Med 12  6E / 6W / Boyer
Med 5   5E / 5W             Med 13  Overflow
Med 6   6E / 6W / Boyer     Med 14  Overflow
Med 7   7E / 7W             Med 15  Overflow
Med 8   8E / 8W

* = IMCU teams (cap: 10)
Room convention: X01-X20 = West, X30-X50 = East
"""
    st.code(ref_text, language=None)

# Main tabs
tab_nights, tab_shuffle = st.tabs(["Overnight Redis", "Monday Shuffle"])

# =============================================================================
# TAB 1: NIGHTS (Overnight Placement)
# =============================================================================

with tab_nights:
    # Create 5 columns: Census + 4 Doctors side by side
    census_col, doc1_col, doc2_col, doc3_col, doc4_col = st.columns([1, 1, 1, 1, 1])

    # Column 1: Team Census (compact layout)
    with census_col:
        st.subheader("Census")
        st.caption("X = closed")

        nights_census = {}
        nights_closed_teams = set()

        for team in ALL_TEAMS:
            label_col, input_col = st.columns([1, 3])
            with label_col:
                imcu = "*" if team in IMCU_TEAMS else ""
                st.markdown(f"<div style='padding-top:8px; text-align:right'>Med {team}{imcu}</div>", unsafe_allow_html=True)
            with input_col:
                value = st.text_input(
                    f"Med {team}",
                    key=f"nights_census_{team}",
                    label_visibility="collapsed"
                )

            if value:
                value_upper = value.strip().upper()
                if value_upper in ('X', 'NA', 'CLOSED', 'N/A', '-'):
                    nights_closed_teams.add(team)
                    nights_census[team] = 0
                else:
                    try:
                        nights_census[team] = int(value)
                    except ValueError:
                        nights_census[team] = 0
            else:
                nights_census[team] = 0

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
                key=f"nights_doc_{i}",
                placeholder="Name",
                label_visibility="collapsed"
            )
            doctor_names.append(name.strip() if name else code)

            patients = st.text_area(
                "Patients",
                key=f"nights_patients_{i}",
                height=400,
                placeholder="310A\n545\n634* (append * for IMCU)\n..." if i == 1 else "312\n545\n7E\n...",
                label_visibility="collapsed"
            )
            doctor_patients.append(patients)

    # Show closed teams
    if nights_closed_teams:
        st.info(f"**Closed teams:** {', '.join(f'Med {t}' for t in sorted(nights_closed_teams))}")

    # Process button
    if st.button("Optimize Placements", type="primary", use_container_width=True, key="nights_optimize"):
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

                is_imcu_override = False
                if location.endswith('*'):
                    is_imcu_override = True
                    location = location[:-1]

                location_key = location.upper()
                if location_key in seen:
                    duplicates += 1
                    continue
                seen.add(location_key)

                if is_imcu_override:
                    floor = 'IMCU'
                else:
                    floor = normalize_floor(location)

                patients.append(Patient(f"Pt{pt_count}", floor, location, admitted_by=doc_name))
                pt_count += 1

        if not patients:
            st.warning("No patients entered. Please enter patient locations above.")
        else:
            assignments = optimize_placements(patients, nights_census, nights_closed_teams)

            final_census = nights_census.copy()
            for a in assignments:
                final_census[a.team] = final_census.get(a.team, 0) + 1

            st.markdown("---")
            st.markdown("<div id='results-section'></div>", unsafe_allow_html=True)
            st.subheader("Results")

            components.html("""
                <script>
                    window.parent.document.getElementById('results-section').scrollIntoView({behavior: 'smooth'});
                </script>
            """, height=0)

            if duplicates > 0:
                st.warning(f"Skipped {duplicates} duplicate(s)")

            geo_count = sum(1 for a in assignments if a.is_geographic)
            total = len(assignments)

            metric_cols = st.columns(4)
            metric_cols[0].metric("Total Patients", total)
            metric_cols[1].metric("Geographic", f"{geo_count} ({100*geo_count//total if total else 0}%)")
            metric_cols[2].metric("Non-Geographic", total - geo_count)
            metric_cols[3].metric("Teams Used", len(set(a.team for a in assignments)))

            res_col1, res_col2, res_col3, res_col4 = st.columns(4)

            with res_col1:
                st.markdown("### Census Summary")
                summary_text = "Team   Start +New =Final\n"
                summary_text += "-" * 26 + "\n"

                for team in ALL_TEAMS:
                    if team in nights_closed_teams:
                        summary_text += f"Med {team:2d}   -- CLOSED\n"
                        continue

                    start = nights_census.get(team, 0)
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

            with res_col2:
                st.markdown("### Assignment List")
                sorted_assignments = sorted(assignments, key=lambda a: a.patient.raw_location)
                assignment_text = ""
                for a in sorted_assignments:
                    assignment_text += f"{a.patient.raw_location:8} → Med {a.team:2d} ({a.patient.admitted_by})\n"
                st.code(assignment_text, language=None)

            with res_col3:
                st.markdown("### By Team")
                by_team = defaultdict(list)
                for a in assignments:
                    by_team[a.team].append(a)

                by_team_text = ""
                for team in ALL_TEAMS:
                    if team in nights_closed_teams:
                        continue
                    team_assignments = by_team[team]
                    if not team_assignments:
                        continue

                    start = nights_census.get(team, 0)
                    final = final_census.get(team, 0)

                    imcu = "*" if team in IMCU_TEAMS else ""
                    by_team_text += f"Med {team}{imcu} ({start}→{final})\n"

                    for a in team_assignments:
                        by_team_text += f"  {a.patient.raw_location} ({a.patient.admitted_by})\n"
                    by_team_text += "\n"

                st.code(by_team_text, language=None)

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
                line_count = len(epic_lines)
                text_height = max(150, line_count * 22 + 50)

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


# =============================================================================
# TAB 2: MONDAY SHUFFLE
# =============================================================================

with tab_shuffle:
    st.markdown("Identify patients on wrong teams for manual reassignment in Epic.")

    # Closed teams input
    shuffle_closed_input = st.text_input(
        "Closed teams (comma-separated, e.g., '14, 15')",
        value="14, 15",
        help="Enter team numbers that are closed",
        key="shuffle_closed"
    )
    shuffle_closed_teams = set()
    if shuffle_closed_input:
        for t in shuffle_closed_input.split(','):
            t = t.strip()
            if t.isdigit():
                shuffle_closed_teams.add(int(t))

    st.markdown("---")

    # Input method tabs
    ocr_tab, manual_tab = st.tabs(["Screenshot OCR", "Manual Entry"])

    # Session state for shuffle
    if 'shuffle_patients' not in st.session_state:
        st.session_state.shuffle_patients = []
    if 'shuffle_uploader_key' not in st.session_state:
        st.session_state.shuffle_uploader_key = 0

    with ocr_tab:
        st.markdown("""
        Upload EPIC screenshots with room numbers and team assignments only. **DO NOT UPLOAD PHI.**

        For best results, crop screenshots to show just the Room and Team columns.
        """)

        if not OCR_AVAILABLE:
            st.error("OCR (pytesseract) not available. Use Manual Entry tab instead.")
        else:
            uploaded_files = st.file_uploader(
                "Upload screenshots",
                type=['png', 'jpg', 'jpeg'],
                accept_multiple_files=True,
                key=f"shuffle_uploader_{st.session_state.shuffle_uploader_key}"
            )

            if uploaded_files:
                process_btn = st.button("Process Screenshots", type="primary", key="shuffle_ocr_btn")
            else:
                process_btn = False

            if process_btn:
                all_pairs = []

                for uploaded_file in uploaded_files:
                    with st.spinner(f"Processing {uploaded_file.name}..."):
                        try:
                            image = Image.open(uploaded_file)

                            # Try EXIF transpose first (handles phone photo orientation)
                            try:
                                image = ImageOps.exif_transpose(image)
                            except Exception:
                                pass

                            # Use rotation-aware OCR that tries all orientations
                            best_pairs, best_text, best_rotation, best_psm = try_all_rotations(image)
                        except Exception as e:
                            st.error(f"Error processing {uploaded_file.name}: {e}")
                            best_pairs, best_text, best_rotation, best_psm = [], "", 0, 6

                    with st.expander(f"Raw OCR from {uploaded_file.name}"):
                        st.code(best_text)
                        rooms_found = re.findall(r'\b(\d{3,4}[A-Z]?)\b', best_text, re.IGNORECASE)
                        teams_found = re.findall(r'Med\s*(\d{1,2})', best_text, re.IGNORECASE)
                        rotation_info = f", rotated {best_rotation}°" if best_rotation != 0 else ""
                        st.caption(f"Debug: {len(rooms_found)} rooms, {len(teams_found)} teams (PSM {best_psm}{rotation_info})")

                    all_pairs.extend(best_pairs)
                    st.info(f"Found {len(best_pairs)} room-team pairs in {uploaded_file.name}")

                seen_rooms = set()
                st.session_state.shuffle_patients = []
                for room, team in all_pairs:
                    room = fix_ocr_room(room)
                    if room is None:
                        continue
                    if room not in seen_rooms:
                        seen_rooms.add(room)
                        floor = normalize_floor(room)
                        st.session_state.shuffle_patients.append(
                            ExistingPatient(room=room, current_team=team, floor=floor)
                        )

                if st.session_state.shuffle_patients:
                    st.success(f"Total: {len(st.session_state.shuffle_patients)} unique patients extracted")

                    with st.expander("Extracted data (verify this is correct)"):
                        extracted_text = ""
                        for p in sorted(st.session_state.shuffle_patients, key=lambda x: x.room):
                            extracted_text += f"{p.room} Med {p.current_team}\n"
                        st.code(extracted_text)
                else:
                    st.warning("No room-team pairs found. Check the raw OCR output above.")

    with manual_tab:
        st.markdown("""
        **Enter room and current team** - one per line.

        Format: `Room Team` (e.g., `304A 1` or `304A Med 1`)
        """)

        paste_input = st.text_area(
            "Room and Team (one per line)",
            height=400,
            placeholder="304A 1\n343B 5\n534 Med 10\n435A 7\n...",
            key="shuffle_manual_input"
        )

        if st.button("Analyze Teams", type="primary", key="shuffle_manual_btn"):
            if not paste_input:
                st.warning("Please enter room and team data first.")
            else:
                seen_rooms = set()
                parse_errors = []
                st.session_state.shuffle_patients = []

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
                    st.session_state.shuffle_patients.append(
                        ExistingPatient(room=room, current_team=team_num, floor=floor)
                    )

                if parse_errors:
                    with st.expander(f"Parse warnings ({len(parse_errors)})"):
                        for err in parse_errors:
                            st.warning(err)

                st.success(f"Parsed {len(st.session_state.shuffle_patients)} patients")

    # Process results if we have patients
    if st.session_state.shuffle_patients:
        all_patients = st.session_state.shuffle_patients
        wrong_team, ok_team = analyze_patients(all_patients, shuffle_closed_teams)

        st.markdown("---")

        header_col1, header_col2 = st.columns([3, 17])
        with header_col1:
            st.markdown('<p style="font-size: 1.5rem; font-weight: 600; margin: 0; display: inline;">Analysis</p>', unsafe_allow_html=True)
        with header_col2:
            if st.button("Clear", key="shuffle_clear_btn", type="secondary"):
                st.session_state.shuffle_patients = []
                st.session_state.shuffle_uploader_key += 1
                st.rerun()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Patients", len(all_patients))
        col2.metric("Need Reassignment", len(wrong_team))
        col3.metric("Team Correct", len(ok_team))

        team_census = defaultdict(int)
        for patient in all_patients:
            team_census[patient.current_team] += 1

        SHUFFLE_IMCU_TARGET = 9
        OVERFLOW_THRESHOLD = 10

        projected_census = dict(team_census)
        recommendations = []

        available_overflow = [t for t in OVERFLOW_TEAMS if t not in shuffle_closed_teams]

        for patient, acceptable in sorted(wrong_team, key=lambda x: x[0].room):
            if acceptable:
                def team_score(t):
                    census = projected_census.get(t, 0)
                    if t in IMCU_TEAMS and census >= SHUFFLE_IMCU_TARGET:
                        return census + 100
                    return census

                min_geo_census = min(projected_census.get(t, 0) for t in acceptable)
                if min_geo_census >= OVERFLOW_THRESHOLD and available_overflow:
                    options = acceptable + available_overflow
                else:
                    options = acceptable

                best_team = min(options, key=team_score)
                projected_census[best_team] = projected_census.get(best_team, 0) + 1
                projected_census[patient.current_team] = projected_census.get(patient.current_team, 0) - 1
                recommendations.append((patient, best_team))
            else:
                if available_overflow:
                    best_team = min(available_overflow, key=lambda t: projected_census.get(t, 0))
                    projected_census[best_team] = projected_census.get(best_team, 0) + 1
                    projected_census[patient.current_team] = projected_census.get(patient.current_team, 0) - 1
                    recommendations.append((patient, best_team))
                else:
                    recommendations.append((patient, None))

        res_col1, res_col2, res_col3 = st.columns(3)

        with res_col1:
            st.markdown("### Census")
            census_text = "Team  Now  +/-  =New\n"
            census_text += "-" * 22 + "\n"
            for team in ALL_TEAMS:
                if team in shuffle_closed_teams:
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

        st.markdown("---")
        st.markdown("### New Assignments by Team")

        team_rosters = defaultdict(list)
        team_leaving = defaultdict(int)

        for patient in ok_team:
            team_rosters[patient.current_team].append((patient.room, "stays"))

        for patient, rec_team in recommendations:
            if rec_team:
                team_rosters[rec_team].append((patient.room, "new"))
                team_leaving[patient.current_team] += 1

        teams_to_show = [t for t in ALL_TEAMS if t not in shuffle_closed_teams]

        for row_start in range(0, len(teams_to_show), 5):
            row_teams = teams_to_show[row_start:row_start + 5]
            cols = st.columns(5)

            for col, team in zip(cols, row_teams):
                with col:
                    roster = team_rosters[team]
                    new_count = sum(1 for _, status in roster if status == "new")
                    old_count = team_leaving[team]
                    imcu = "*" if team in IMCU_TEAMS else ""
                    floors = TEAM_FLOORS_STR.get(team, "")
                    proj = projected_census.get(team, 0)

                    roster_text = f"Med {team}{imcu} ({proj})\n"
                    roster_text += f"{floors}\n"
                    roster_text += "-" * 14 + "\n"

                    if roster or old_count > 0:
                        for room, status in sorted(roster):
                            marker = ">" if status == "new" else " "
                            roster_text += f"{marker} {room}\n"
                        roster_text += f"\n-{old_count} old, +{new_count} new"
                    else:
                        roster_text += "(no changes)"

                    st.code(roster_text, language=None)
