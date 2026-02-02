#!/usr/bin/env python3
"""
Geographic Patient Placement Optimizer

Recommends optimal team assignments for patients based on:
1. Geographic match (patient floor → team's assigned floors)
2. Census balancing (distribute evenly across teams)

Use cases:
- Monday redistribution before new attendings start Tuesday
- Morning T-list distribution of overnight admissions
"""

import re
from dataclasses import dataclass
from typing import Optional
from collections import defaultdict

# =============================================================================
# GEOGRAPHIC MAPPINGS
# =============================================================================

# Floor → Teams that cover that floor (direct care teams)
# Room convention: X01-X20 = West, X30-X50 = East
FLOOR_TO_TEAMS = {
    '3W': [1, 2, 3],
    '3E': [1, 2, 3],
    '4W': [4, 11],
    '4E': [4, 11],
    '5W': [5, 10],
    '5E': [5, 10],
    '6W': [6, 12],
    '6E': [6, 12],
    '7W': [7, 9],
    '7E': [7, 9],
    '8E': [8, 13],
    '8W': [8, 13],      # If 8W exists, same teams as 8E
    # Special units
    'IMCU': [1, 2, 3],  # IMCU is on 3E
    'BOYER': [12, 6],   # Boyer: preferentially Med 12, overflow to Med 6, then anywhere
}

# Team → Floors (for display)
TEAM_FLOORS = {
    1: ['3W', '3E', 'IMCU'],
    2: ['3W', '3E', 'IMCU'],
    3: ['3W', '3E', 'IMCU'],
    4: ['4E', '4W'],
    5: ['5E', '5W'],
    6: ['6E', '6W'],
    7: ['7E', '7W'],
    8: ['8E'],
    9: ['7E', '7W'],
    10: ['5E', '5W'],
    11: ['4E', '4W'],
    12: ['6E', '6W', 'Boyer'],
    13: ['8E'],
    14: ['Overflow'],
    15: ['Overflow'],
}

# All valid teams
ALL_TEAMS = list(range(1, 16))  # Med 1-15
OVERFLOW_TEAMS = [14, 15]  # Only used when regular teams are full

# IMCU teams (Med 1-3) - lower cap due to higher acuity
IMCU_TEAMS = [1, 2, 3]
IMCU_CAP = 10  # Hard cap for IMCU teams

# Soft cap for regular teams - try not to exceed
SOFT_CAP = 14  # Avoid loading teams above this

# Scoring weights (lower score = better)
# score = (census * CENSUS_WEIGHT) + (redis * REDIS_WEIGHT) + penalty
# penalty = 0 if geographic, GEO_PENALTY if not (or IMCU_PENALTY for 3W/IMCU patients)
CENSUS_WEIGHT = 1.0    # Weight for current census
REDIS_WEIGHT = 1.0     # Weight for new patients (redis) tonight
GEO_PENALTY = 3.0      # Penalty for non-geographic placement
IMCU_PENALTY = 10.0    # Higher penalty for 3W/IMCU patients going off-floor (they NEED Med 1-3)

# =============================================================================
# DATA STRUCTURES
# =============================================================================

@dataclass
class Patient:
    identifier: str  # Room/bed number or name
    floor: str       # Normalized floor (e.g., "5E")
    raw_location: str  # Original input


@dataclass
class Assignment:
    patient: Patient
    team: int
    is_geographic: bool
    reason: str


# =============================================================================
# FLOOR PARSING
# =============================================================================

def normalize_floor(location: str) -> Optional[str]:
    """
    Parse a location string and extract the floor.

    Room numbering convention:
        X01-X20 = Floor X West
        X30-X50 = Floor X East
        750+, 850+, 9W = Boyer wing → Med 12
        Suffix A/B ignored (312A = 312B = same room)

    Examples:
        "312A" → "3W" (room 12 on floor 3 = West)
        "545B" → "5E" (room 45 on floor 5 = East)
        "877" → "BOYER" (Boyer wing)
        "IMCU" → "IMCU" (Med 1, 2, 3)
        "ED" → None (can go anywhere)
    """
    original = location.strip().upper()

    # Special locations - check BEFORE modifying the string
    # IMCU is on 3E, maps to Med 1, 2, 3
    if 'IMCU' in original:
        return 'IMCU'

    # Overnight recovery → Med 12 (Boyer)
    if 'OVERNIGHT' in original or 'ONR' in original or 'RECOVERY' in original:
        return 'BOYER'

    # ED patients - can go to any team (lowest census)
    if any(x in original for x in ['ED', 'EMERGENCY', 'ER ']):
        return None  # Will be assigned to lowest census team

    # Boyer explicitly mentioned
    if 'BOYER' in original:
        return 'BOYER'

    # Now safe to remove bed suffix (A, B, etc.) for room number parsing
    location = re.sub(r'[A-Z]$', '', original)

    # Pattern: explicit floor + direction (e.g., "5E", "7W", "3 East", "5E-512")
    match = re.search(r'(\d+)\s*([EW]|EAST|WEST)', location)
    if match:
        floor_num = match.group(1)
        direction = match.group(2)[0]  # First letter: E or W
        # 9W is Boyer
        if floor_num == '9' and direction == 'W':
            return 'BOYER'
        return f"{floor_num}{direction}"

    # Pattern: room number (e.g., "312", "545", "877")
    match = re.search(r'\b(\d)(\d{2})\b', location)
    if match:
        floor_num = match.group(1)
        room_num = int(match.group(2))

        # Boyer wing: 750+ on floor 7, 850+ on floor 8
        if floor_num == '7' and room_num >= 50:
            return 'BOYER'
        if floor_num == '8' and room_num >= 50:
            return 'BOYER'
        # 9xx rooms are Boyer
        if floor_num == '9':
            return 'BOYER'

        # Standard rooms
        if 1 <= room_num <= 20:
            return f"{floor_num}W"
        elif 30 <= room_num < 50:
            return f"{floor_num}E"
        else:
            # Room number outside normal range
            return f"{floor_num}?"

    # Pattern: "Floor 5 East" or similar
    match = re.search(r'FLOOR\s*(\d+)\s*(EAST|WEST|E|W)', location)
    if match:
        floor_num = match.group(1)
        direction = match.group(2)[0]
        return f"{floor_num}{direction}"

    return None


def get_geographic_teams(floor: str) -> list[int]:
    """Get teams that cover a given floor."""
    if not floor:
        return []

    if floor in FLOOR_TO_TEAMS:
        return FLOOR_TO_TEAMS[floor]

    # Handle ambiguous floor (e.g., "5?" when room number was outside 01-20, 30-50)
    if floor.endswith('?'):
        floor_num = floor[:-1]
        # Try both East and West for that floor
        east_teams = FLOOR_TO_TEAMS.get(f"{floor_num}E", [])
        west_teams = FLOOR_TO_TEAMS.get(f"{floor_num}W", [])
        # Return union of both (they're usually the same anyway)
        return list(set(east_teams + west_teams))

    return []


# =============================================================================
# PLACEMENT ALGORITHM
# =============================================================================

def get_patient_priority(patient: Patient) -> tuple:
    """
    Get sort priority for a patient. Lower = processed first.

    Order:
    0. 3W/IMCU/* patients (NEED Med 1-3, can take 10th slot)
    1. Outliers (ED, unknown) - balance census first
    2. 3E + other geographic patients (3E prefers 1-3 but can't take 10th slot)
    """
    floor = patient.floor
    if not floor:
        # Outliers go BEFORE regular geo patients to balance census
        return (1, 'ZZZ', patient.identifier)

    geo_teams = get_geographic_teams(floor)
    if not geo_teams:
        # Unknown floor = outlier
        return (1, floor, patient.identifier)

    # 3W and IMCU (includes * patients) - NEED to go to Med 1-3
    if floor in ['3W', 'IMCU']:
        return (0, floor, patient.identifier)

    # 3E and other geographic floors - process AFTER outliers
    return (2, floor, patient.identifier)


def optimize_placements(
    patients: list[Patient],
    current_census: dict[int, int],
    closed_teams: set[int] = None
) -> list[Assignment]:
    """
    Assign patients to teams using weighted scoring for census, redis, and geography.

    Scoring: score = (census * CENSUS_WEIGHT) + (redis * REDIS_WEIGHT) + penalty
    - penalty = 0 if geographic
    - penalty = IMCU_PENALTY (10) for 3W/IMCU patients going off-floor
    - penalty = GEO_PENALTY (3) for other patients going off-floor

    Lower score = better assignment.

    Patient order: 3W/IMCU/* (NEED) → outliers (balance) → 3E + other geo.
    IMCU teams at census 9 reserve the 10th slot for IMCU patients only.

    Hard constraints:
    1. Never assign to closed teams
    2. Med 1-3 (IMCU) hard cap at 10 patients total
    3. Prefer teams under soft cap (14) when possible
    """
    assignments = []
    census = current_census.copy()
    new_assignments = {t: 0 for t in ALL_TEAMS}  # Track NEW patients (redis) per team
    closed_teams = closed_teams or set()

    # Get list of open teams
    open_teams = [t for t in ALL_TEAMS if t not in closed_teams]
    regular_open_teams = [t for t in open_teams if t not in OVERFLOW_TEAMS]

    # Sort patients: 3W/IMCU/* first, then outliers, then 3E + other geo
    patients_sorted = sorted(patients, key=get_patient_priority)

    for patient in patients_sorted:
        geo_teams = set(get_geographic_teams(patient.floor) if patient.floor else [])
        geo_teams -= closed_teams  # Remove closed teams

        # Determine penalty for this patient
        is_imcu_patient = patient.floor in ['3W', 'IMCU']
        non_geo_penalty = IMCU_PENALTY if is_imcu_patient else GEO_PENALTY

        def team_score(t: int) -> tuple[float, int, int]:
            """
            Score a team for this patient. Lower = better.
            Returns (score, census, is_not_geo) for sorting.
            """
            c = census.get(t, 0)
            r = new_assignments[t]
            is_geo = t in geo_teams

            # Score = census + redis + penalty
            score = (c * CENSUS_WEIGHT) + (r * REDIS_WEIGHT)
            if not is_geo:
                score += non_geo_penalty

            # Penalty for piling on: if this team has 2+ redis and others have 0
            if r >= 2:
                teams_with_zero = [tm for tm in regular_open_teams if new_assignments[tm] == 0]
                if teams_with_zero:
                    score += 2.0  # Discourage giving 3rd+ patient when others have none

            # Return tuple for stable sorting: (score, current_census, not_geo)
            return (score, c, 0 if is_geo else 1)

        def is_eligible(t: int) -> bool:
            """Check if team can accept patients (hard constraints)."""
            if t in IMCU_TEAMS:
                current = census.get(t, 0)
                # Hard cap at 10
                if current >= IMCU_CAP:
                    return False
                # Reserve 10th slot for IMCU patients only (3W, IMCU, or * suffix)
                if current == IMCU_CAP - 1 and patient.floor not in ['3W', 'IMCU']:
                    return False
            return True

        # Get eligible teams, preferring regular teams over overflow
        # First try: regular teams under soft cap
        candidates = [t for t in regular_open_teams if is_eligible(t) and census.get(t, 0) < SOFT_CAP]

        # If no regular teams under soft cap, try overflow teams
        if not candidates:
            overflow_open = [t for t in OVERFLOW_TEAMS if t in open_teams]
            candidates = [t for t in overflow_open if is_eligible(t) and census.get(t, 0) < SOFT_CAP]

        # If still none, allow regular teams over soft cap
        if not candidates:
            candidates = [t for t in regular_open_teams if is_eligible(t)]

        # Last resort: any open team that's eligible
        if not candidates:
            candidates = [t for t in open_teams if is_eligible(t)]

        if not candidates:
            print(f"WARNING: No eligible teams for {patient.raw_location}")
            continue

        # Pick best team by score
        best_team = min(candidates, key=team_score)
        is_geo = best_team in geo_teams

        # Build reason string
        score_val = team_score(best_team)[0]
        if is_geo:
            reason = f"Geographic ({patient.floor} → Med {best_team}, score={score_val:.1f})"
        else:
            if patient.floor:
                reason = f"Balance ({patient.floor} → Med {best_team}, score={score_val:.1f})"
            else:
                reason = f"No floor specified (Med {best_team}, score={score_val:.1f})"

        # Update tracking
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
# INTERACTIVE INTERFACE
# =============================================================================

def get_census_input() -> tuple[dict[int, int], set[int]]:
    """Interactively get current census for each team.

    Returns:
        Tuple of (census dict, set of closed teams)
    """
    print("\n" + "=" * 60)
    print("STEP 1: Enter Current Team Census")
    print("=" * 60)
    print("Enter census for each team (Enter=0, NA/X=closed)")
    print()

    census = {}
    closed_teams = set()

    for team in ALL_TEAMS:
        floors = ', '.join(TEAM_FLOORS[team])
        imcu_marker = " [IMCU]" if team in IMCU_TEAMS else ""
        while True:
            value = input(f"  Med {team:2d} ({floors}){imcu_marker}: ").strip().upper()

            # Check for closed/NA
            if value in ('NA', 'X', 'CLOSED', 'N/A', '-'):
                census[team] = 0
                closed_teams.add(team)
                print(f"         → CLOSED")
                break

            # Check for number
            try:
                census[team] = int(value) if value else 0
                break
            except ValueError:
                print("    Enter a number, or NA/X if closed")

    if closed_teams:
        print(f"\n  Closed teams: {', '.join(f'Med {t}' for t in sorted(closed_teams))}")

    return census, closed_teams


def get_patients_input() -> list[Patient]:
    """Interactively get list of patients to place."""
    print("\n" + "=" * 60)
    print("STEP 2: Enter Patients to Place")
    print("=" * 60)
    print("Enter patient locations one per line.")
    print("Format: room number or floor (e.g., '512', '5E-512', '7W')")
    print("Append * for IMCU placement (e.g., '614*' = room 614 going to IMCU)")
    print("Type 'done' when finished.")
    print()

    patients = []
    seen_locations = set()  # Track duplicates
    count = 1
    duplicates = 0

    while True:
        location = input(f"  Patient {count}: ").strip()

        if location.lower() == 'done':
            break

        if not location:
            continue

        # Check for IMCU override (trailing *)
        is_imcu_override = False
        if location.endswith('*'):
            is_imcu_override = True
            location = location[:-1]

        # Normalize for duplicate detection (uppercase, remove trailing letters)
        location_key = location.upper().strip()

        # Check for duplicate
        if location_key in seen_locations:
            print(f"         → DUPLICATE (skipping)")
            duplicates += 1
            continue

        seen_locations.add(location_key)

        # Set floor - IMCU override takes precedence
        if is_imcu_override:
            floor = 'IMCU'
        else:
            floor = normalize_floor(location)

        patients.append(Patient(
            identifier=f"Pt{count}",
            floor=floor,
            raw_location=location + ('*' if is_imcu_override else '')
        ))

        if floor:
            geo_teams = get_geographic_teams(floor)
            if geo_teams:
                teams_str = ', '.join(f"Med {t}" for t in geo_teams)
                imcu_note = " [IMCU priority]" if is_imcu_override else ""
                print(f"         → {floor} (geographic: {teams_str}){imcu_note}")
            else:
                print(f"         → {floor} (any team)")
        else:
            print(f"         → Any team (ED/unknown)")

        count += 1

    if duplicates > 0:
        print(f"\n  ⚠️  Skipped {duplicates} duplicate(s)")

    return patients


def display_results(assignments: list[Assignment], final_census: dict[int, int], starting_census: dict[int, int], closed_teams: set[int] = None):
    """Display the recommended assignments."""
    closed_teams = closed_teams or set()
    print("\n" + "=" * 60)
    print("RECOMMENDED ASSIGNMENTS")
    print("=" * 60)

    # Group by team
    by_team = defaultdict(list)
    for a in assignments:
        by_team[a.team].append(a)

    geo_count = sum(1 for a in assignments if a.is_geographic)
    total = len(assignments)

    print(f"\nTotal patients to place: {total}")
    print(f"Geographic placements: {geo_count} ({100*geo_count//total if total else 0}%)")
    print(f"Non-geographic: {total - geo_count}")

    print("\n" + "-" * 60)
    print("BY TEAM:")
    print("-" * 60)

    for team in ALL_TEAMS:
        team_assignments = by_team[team]
        new_count = len(team_assignments)
        if new_count > 0 or final_census.get(team, 0) > 0:
            floors = ', '.join(TEAM_FLOORS[team])
            imcu_marker = " [IMCU cap:10]" if team in IMCU_TEAMS else ""
            start = starting_census.get(team, 0)
            final = final_census.get(team, 0)
            print(f"\nMed {team} ({floors}){imcu_marker}")
            print(f"  Census: {start} → {final} (+{new_count} new)")
            for a in team_assignments:
                geo_marker = "✓" if a.is_geographic else "✗"
                print(f"    {geo_marker} {a.patient.raw_location:15} ({a.patient.floor or '?'})")

    print("\n" + "-" * 60)
    print("ASSIGNMENT LIST (copy/paste ready):")
    print("-" * 60)
    print()

    for a in assignments:
        geo_marker = "GEO" if a.is_geographic else "   "
        print(f"  {a.patient.raw_location:15} → Med {a.team:2d}  {geo_marker}")

    print("\n" + "-" * 60)
    print("FINAL CENSUS SUMMARY:")
    print("-" * 60)
    print()
    print("  Team    Start  +New  =Final")
    print("  " + "-" * 30)

    for team in ALL_TEAMS:
        if team in closed_teams:
            imcu = "*" if team in IMCU_TEAMS else " "
            print(f"  Med {team:2d}{imcu}   --   CLOSED")
            continue

        start = starting_census.get(team, 0)
        final = final_census.get(team, 0)
        new = final - start
        bar = "█" * min(final, 20)  # Cap bar length at 20
        if final > 20:
            bar += "+"
        imcu = "*" if team in IMCU_TEAMS else " "
        if team in IMCU_TEAMS and final >= IMCU_CAP:
            cap_warning = " ⚠️ AT CAP"
        elif team not in IMCU_TEAMS and final >= SOFT_CAP:
            cap_warning = " ⚠️ HIGH"
        else:
            cap_warning = ""
        print(f"  Med {team:2d}{imcu}  {start:3d}   +{new:2d}   ={final:3d}  {bar}{cap_warning}")

    print()
    print("  * = IMCU team (hard cap: 10)")
    print(f"  Other teams soft cap: {SOFT_CAP}")


def run_interactive():
    """Run the interactive placement optimizer."""
    print("=" * 60)
    print("GEOGRAPHIC PATIENT PLACEMENT OPTIMIZER")
    print("=" * 60)
    print()
    print("Scoring: score = census + redis + penalty")
    print(f"  CENSUS_WEIGHT = {CENSUS_WEIGHT}")
    print(f"  REDIS_WEIGHT  = {REDIS_WEIGHT}")
    print(f"  GEO_PENALTY   = {GEO_PENALTY} (regular floors)")
    print(f"  IMCU_PENALTY  = {IMCU_PENALTY} (3W/IMCU patients)")
    print()
    print("Patient order: 3W/IMCU/* (NEED) → outliers → 3E + other geo")
    print("IMCU teams at 9 reserve 10th slot for IMCU patients only.")
    print("Lower score wins. 3W/IMCU patients strongly prefer Med 1-3.")

    # Get inputs
    census, closed_teams = get_census_input()
    patients = get_patients_input()

    if not patients:
        print("\nNo patients entered. Exiting.")
        return

    # Run optimization
    print("\n" + "=" * 60)
    print("OPTIMIZING...")
    print("=" * 60)

    assignments = optimize_placements(patients, census, closed_teams)

    # Calculate final census
    final_census = census.copy()
    for a in assignments:
        final_census[a.team] = final_census.get(a.team, 0) + 1

    # Display results
    display_results(assignments, final_census, census, closed_teams)


# =============================================================================
# QUICK MODE (for T-list distribution)
# =============================================================================

def run_quick_mode():
    """Quick mode - just enter locations, assume empty census."""
    print("=" * 60)
    print("QUICK MODE - T-List Distribution")
    print("=" * 60)
    print()
    print("Enter patient locations (one per line, 'done' to finish):")
    print()

    patients = []
    count = 1

    while True:
        location = input(f"  {count}: ").strip()
        if location.lower() == 'done':
            break
        if not location:
            continue

        floor = normalize_floor(location)
        patients.append(Patient(f"Pt{count}", floor, location))
        count += 1

    if not patients:
        return

    # Assume starting census of 0 for quick distribution, no closed teams
    census = {t: 0 for t in ALL_TEAMS}
    closed_teams = set()
    assignments = optimize_placements(patients, census, closed_teams)

    final_census = {t: 0 for t in ALL_TEAMS}
    for a in assignments:
        final_census[a.team] += 1

    display_results(assignments, final_census, census, closed_teams)


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == '--quick':
        run_quick_mode()
    else:
        run_interactive()
