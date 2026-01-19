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

# Equity rules
MAX_NEW_BEFORE_SPREAD = 3  # Don't give team more than 3 new patients if others have capacity
MAX_CENSUS_GAP = 4  # If assigning to team would create gap > this vs lowest team, prefer balance

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

def optimize_placements(
    patients: list[Patient],
    current_census: dict[int, int],
    closed_teams: set[int] = None
) -> list[Assignment]:
    """
    Assign patients to teams optimizing for geography, balance, and equity.

    Rules:
    1. Never assign to closed teams
    2. Med 1-3 (IMCU) hard cap at 10 patients total
    3. Soft cap of 14 for all other teams - avoid if possible
    4. Don't give a team more than 3 new patients if other geographic teams have capacity
    5. Among valid teams, pick the one with lowest census
    6. More aggressive balancing - prefer low census teams even over geography
    """
    assignments = []
    census = current_census.copy()
    new_assignments = {t: 0 for t in ALL_TEAMS}  # Track NEW patients per team
    closed_teams = closed_teams or set()

    # Get list of open teams (excluding overflow teams for regular assignments)
    open_teams = [t for t in ALL_TEAMS if t not in closed_teams]
    regular_open_teams = [t for t in open_teams if t not in OVERFLOW_TEAMS]

    # Sort patients by floor to group geographic placements
    patients_sorted = sorted(patients, key=lambda p: (p.floor or 'ZZZ', p.identifier))

    for patient in patients_sorted:
        geo_teams_raw = get_geographic_teams(patient.floor) if patient.floor else []
        # Filter out closed teams from geographic options
        geo_teams = [t for t in geo_teams_raw if t not in closed_teams]
        best_team = None
        is_geo = False
        reason = ""

        if geo_teams:
            # Filter out teams that are at capacity or over limits
            valid_geo_teams = []
            for t in geo_teams:
                # Check IMCU hard cap
                if t in IMCU_TEAMS and census.get(t, 0) >= IMCU_CAP:
                    continue
                # Check soft cap for non-IMCU (try to avoid, but not absolute)
                if t not in IMCU_TEAMS and census.get(t, 0) >= SOFT_CAP:
                    # Only skip if there are other geo teams under soft cap
                    others_under_cap = any(
                        census.get(other, 0) < SOFT_CAP
                        and (other not in IMCU_TEAMS or census.get(other, 0) < IMCU_CAP)
                        for other in geo_teams if other != t
                    )
                    if others_under_cap:
                        continue
                # Check equity - if team already has 3+ new, see if others have fewer
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
                # Pick geographic team with lowest census
                best_geo_team = min(valid_geo_teams, key=lambda t: census.get(t, 0))
                best_geo_census = census.get(best_geo_team, 0)

                # Check if this would create too big a gap vs lowest census REGULAR team
                # (Don't compare against overflow teams 14/15)
                non_imcu_regular = [t for t in regular_open_teams if t not in IMCU_TEAMS]
                if non_imcu_regular:
                    lowest_census_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
                    lowest_census = census.get(lowest_census_team, 0)

                    # If geo team would be way higher than lowest, prefer balance
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
                # All geo teams at capacity/limits - check if we should go non-geo instead
                # Find best geo team (least loaded, respecting hard caps)
                available_geo = [t for t in geo_teams if t not in IMCU_TEAMS or census.get(t, 0) < IMCU_CAP]

                if available_geo:
                    best_geo = min(available_geo, key=lambda t: census.get(t, 0))
                    best_geo_census = census.get(best_geo, 0)

                    # Find best non-geo, non-IMCU REGULAR team (exclude overflow teams from balance comparison)
                    non_imcu_non_geo = [t for t in regular_open_teams if t not in IMCU_TEAMS and t not in geo_teams]
                    if non_imcu_non_geo:
                        best_non_geo = min(non_imcu_non_geo, key=lambda t: census.get(t, 0))
                        best_non_geo_census = census.get(best_non_geo, 0)

                        # If geo team is way higher than non-geo, prefer balance over geography
                        # "Way higher" = 3+ more patients
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
            # No geographic match or all geo teams full/over cap
            # For Boyer overflow, any team is fine - but prefer REGULAR teams over overflow (14, 15)
            non_imcu_regular = [t for t in regular_open_teams if t not in IMCU_TEAMS]
            non_imcu_overflow = [t for t in OVERFLOW_TEAMS if t in open_teams and t not in IMCU_TEAMS]

            # First try regular teams under soft cap
            regular_under_cap = [t for t in non_imcu_regular if census.get(t, 0) < SOFT_CAP]
            if regular_under_cap:
                best_team = min(regular_under_cap, key=lambda t: census.get(t, 0))
            elif non_imcu_overflow:
                # All regular teams at/over soft cap - NOW use overflow teams
                overflow_under_cap = [t for t in non_imcu_overflow if census.get(t, 0) < SOFT_CAP]
                if overflow_under_cap:
                    best_team = min(overflow_under_cap, key=lambda t: census.get(t, 0))
                elif non_imcu_regular:
                    # Overflow also at cap, go back to regular teams
                    best_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
                else:
                    best_team = min(non_imcu_overflow, key=lambda t: census.get(t, 0))
            elif non_imcu_regular:
                # No overflow teams open, use regular even if over soft cap
                best_team = min(non_imcu_regular, key=lambda t: census.get(t, 0))
            elif open_teams:
                # All non-IMCU teams closed, use any open team
                best_team = min(open_teams, key=lambda t: census.get(t, 0))
            else:
                # Shouldn't happen - no teams open at all
                print(f"WARNING: No open teams available for {patient.raw_location}")
                continue

            is_geo = False
            if patient.floor == 'BOYER':
                reason = f"Boyer overflow (Med 12 full), lowest census"
            elif patient.floor:
                reason = f"No geographic capacity for {patient.floor}, lowest census"
            else:
                reason = f"No floor specified, lowest census"

        # Update census and new assignment count
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

        # Normalize for duplicate detection (uppercase, remove trailing letters)
        location_key = location.upper().strip()

        # Check for duplicate
        if location_key in seen_locations:
            print(f"         → DUPLICATE (skipping)")
            duplicates += 1
            continue

        seen_locations.add(location_key)

        floor = normalize_floor(location)
        patients.append(Patient(
            identifier=f"Pt{count}",
            floor=floor,
            raw_location=location
        ))

        if floor:
            geo_teams = get_geographic_teams(floor)
            if geo_teams:
                teams_str = ', '.join(f"Med {t}" for t in geo_teams)
                print(f"         → {floor} (geographic: {teams_str})")
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
    print("This tool recommends optimal team assignments based on:")
    print("  • Geographic match (patient floor → team floors)")
    print("  • Census balancing (distribute evenly)")

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
