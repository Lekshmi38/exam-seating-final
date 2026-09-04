import math
import random
import copy
import re
import json
import os
from collections import defaultdict, Counter
import pandas as pd

# Import the modules
from constraint_handler import ConstraintHandler
from rebalancer import Rebalancer

# ===================== CONFIGURATION =====================
BLOCK_ORDER = ["Col 1", "Col 2", "Col 3", "Col 4", "Col 5"]

# Map column names to display names expected by template
BLOCK_MAPPING = {
    "Col 1": "Left1",
    "Col 2": "Left3",
    "Col 3": "Middle2",
    "Col 4": "Right1",
    "Col 5": "Right3"
}

COL_CAPACITY = {
    "Col 1": 7,
    "Col 2": 7,
    "Col 3": 6,
    "Col 4": 7,
    "Col 5": 7
}

ROOM_MAX_CAP = 34  # 7+7+6+7+7 = 34
MAX_SUBJ_PER_COL = 2

# Initialize modules
constraint_handler = ConstraintHandler()
rebalancer = Rebalancer(constraint_handler)


# ===================== DATA PROCESSING =====================
def extract_elective_counts(data):
    """Extract subject-wise student data - for open electives, key is just subject name"""
    subj_map = defaultdict(list)
    counts = Counter()
    dept_map = defaultdict(set)  # Track departments per subject

    for roll, branch, subject in data:
        # Clean subject name - remove course code in parentheses for grouping
        clean_subject = re.sub(r'\s*\([^)]*\)', '', subject).strip()
        key = clean_subject
        subj_map[key].append({"roll": roll, "branch": branch, "subj": key, "full_subj": subject})
        counts[key] += 1
        dept_map[key].add(branch)

    # Sort students within each subject by roll number
    for key in subj_map:
        subj_map[key].sort(key=lambda x: x['roll'])

    return subj_map, dict(counts), dict(dept_map)


def check_adjacent_constraints(room_data, col_name, student):
    """
    Check if placing student in column violates adjacent constraints:
    1. No same subject in adjacent column
    2. No same department in adjacent column
    """
    col_idx = BLOCK_ORDER.index(col_name)

    # Check left adjacent column
    if col_idx > 0:
        left_col = BLOCK_ORDER[col_idx - 1]
        for existing_student in room_data[left_col]:
            if existing_student['subj'] == student['subj']:
                return False, f"Same subject in adjacent left column"
            if existing_student['branch'] == student['branch']:
                return False, f"Same department in adjacent left column"

    # Check right adjacent column
    if col_idx < len(BLOCK_ORDER) - 1:
        right_col = BLOCK_ORDER[col_idx + 1]
        for existing_student in room_data[right_col]:
            if existing_student['subj'] == student['subj']:
                return False, f"Same subject in adjacent right column"
            if existing_student['branch'] == student['branch']:
                return False, f"Same department in adjacent right column"

    return True, "OK"


def get_column_priority(room_data):
    """Get columns sorted by fill percentage (lowest first) for uniform distribution"""
    col_fill = {}
    for blk in BLOCK_ORDER:
        fill_pct = len(room_data[blk]) / COL_CAPACITY[blk]
        col_fill[blk] = fill_pct

    # Return columns sorted by fill percentage (lowest first)
    return sorted(BLOCK_ORDER, key=lambda x: col_fill[x])


# ===================== ALLOCATION ENGINE =====================
def generate_allocation(subj_map_orig, elective_counts, dept_map, num_rooms):
    """Generate room allocation for open electives with uniform distribution"""
    # Create a fresh deep copy for this allocation attempt
    working_subj_map = copy.deepcopy(subj_map_orig)

    # Create a list of subjects sorted by count (largest first)
    subjects = sorted(
        elective_counts.keys(),
        key=lambda x: elective_counts[x],
        reverse=True
    )

    # Initialize rooms with empty blocks
    rooms = {
        f"Room{i}": {blk: [] for blk in BLOCK_ORDER}
        for i in range(1, num_rooms + 1)
    }

    # Track remaining counts
    remaining_counts = elective_counts.copy()

    # Track subjects per column
    room_col_subjects = {
        r_name: {blk: set() for blk in BLOCK_ORDER}
        for r_name in rooms
    }

    # Track department distribution per column
    room_col_depts = {
        r_name: {blk: defaultdict(int) for blk in BLOCK_ORDER}
        for r_name in rooms
    }

    # Target distribution - try to fill columns evenly
    total_students = sum(elective_counts.values())
    target_per_room = math.ceil(total_students / num_rooms)

    # First pass: Distribute large subjects evenly
    for subject in subjects:
        if remaining_counts[subject] <= 0:
            continue

        # Get sample departments for this subject
        subject_depts = set()
        if subject in working_subj_map and working_subj_map[subject]:
            subject_depts = {s['branch'] for s in working_subj_map[subject][:10]}

        # Try to place in rooms with lowest occupancy first
        while remaining_counts[subject] > 0:
            placed_in_this_round = False

            # Sort rooms by total occupancy (lowest first)
            room_occupancy = []
            for r_name in rooms:
                total_occ = sum(len(col) for col in rooms[r_name].values())
                room_occupancy.append((r_name, total_occ))

            room_occupancy.sort(key=lambda x: x[1])

            for r_name, current_occ in room_occupancy:
                if remaining_counts[subject] <= 0:
                    break

                # Skip if room is already at target + buffer
                if current_occ >= target_per_room + 3:
                    continue

                # Get columns sorted by fill percentage
                columns = get_column_priority(rooms[r_name])

                for blk in columns:
                    if remaining_counts[subject] <= 0:
                        break

                    col = rooms[r_name][blk]
                    used = len(col)
                    cap = COL_CAPACITY[blk]

                    if used >= cap:
                        continue

                    # Check if column already has max subjects
                    if len(room_col_subjects[r_name][blk]) >= MAX_SUBJ_PER_COL:
                        continue

                    # Check if this subject is already in this column
                    if subject in room_col_subjects[r_name][blk]:
                        continue

                    # Find suitable students from this subject
                    suitable_students = []
                    students_to_check = working_subj_map[subject][:min(20, len(working_subj_map[subject]))]

                    for student in students_to_check:
                        # Check adjacent constraints
                        is_adj_safe, _ = check_adjacent_constraints(rooms[r_name], blk, student)
                        if is_adj_safe:
                            suitable_students.append(student)

                    if not suitable_students:
                        continue

                    # Calculate how many we can place
                    free = cap - used
                    take = min(free, remaining_counts[subject], len(suitable_students))

                    if take <= 0:
                        continue

                    # Take students
                    students_to_add = suitable_students[:take]

                    # Add to room
                    rooms[r_name][blk].extend(students_to_add)

                    # Update tracking
                    room_col_subjects[r_name][blk].add(subject)
                    for student in students_to_add:
                        room_col_depts[r_name][blk][student['branch']] += 1

                    # Remove from working map
                    for student in students_to_add:
                        if student in working_subj_map[subject]:
                            working_subj_map[subject].remove(student)

                    remaining_counts[subject] -= len(students_to_add)
                    placed_in_this_round = True

            if not placed_in_this_round:
                break

    # Second pass: Fill remaining students with focus on column balancing
    for subject in subjects:
        if remaining_counts[subject] <= 0:
            continue

        while remaining_counts[subject] > 0:
            placed_in_this_round = False

            for r_name in rooms:
                if remaining_counts[subject] <= 0:
                    break

                # Get columns with lowest fill percentage first
                columns = get_column_priority(rooms[r_name])

                for blk in columns:
                    if remaining_counts[subject] <= 0:
                        break

                    col = rooms[r_name][blk]
                    used = len(col)
                    cap = COL_CAPACITY[blk]

                    if used >= cap:
                        continue

                    # Check if we can add more subjects to this column
                    current_subjects = room_col_subjects[r_name][blk]

                    # If column already has MAX_SUBJ_PER_COL, skip
                    if len(current_subjects) >= MAX_SUBJ_PER_COL and subject not in current_subjects:
                        continue

                    # Find suitable students
                    suitable_students = []
                    students_to_check = working_subj_map[subject][:min(30, len(working_subj_map[subject]))]

                    for student in students_to_check:
                        # Check adjacent constraints
                        is_adj_safe, _ = check_adjacent_constraints(rooms[r_name], blk, student)
                        if is_adj_safe:
                            suitable_students.append(student)

                    if not suitable_students:
                        continue

                    # Calculate how many we can place
                    free = cap - used
                    take = min(free, remaining_counts[subject], len(suitable_students))

                    if take <= 0:
                        continue

                    # Take students
                    students_to_add = suitable_students[:take]

                    # Add to room
                    rooms[r_name][blk].extend(students_to_add)

                    # Update tracking
                    room_col_subjects[r_name][blk].add(subject)
                    for student in students_to_add:
                        room_col_depts[r_name][blk][student['branch']] += 1

                    # Remove from working map
                    for student in students_to_add:
                        if student in working_subj_map[subject]:
                            working_subj_map[subject].remove(student)

                    remaining_counts[subject] -= len(students_to_add)
                    placed_in_this_round = True

            if not placed_in_this_round:
                break

    # Third pass: Balance columns within each room
    for r_name in rooms:
        # Calculate target per column based on room total
        room_total = sum(len(col) for col in rooms[r_name].values())
        if room_total == 0:
            continue

        # Try to fill emptier columns
        for _ in range(50):  # Limited attempts
            moved = False

            # Find columns with imbalance
            col_fill = {}
            for blk in BLOCK_ORDER:
                col_fill[blk] = len(rooms[r_name][blk])

            max_col = max(col_fill, key=lambda x: col_fill[x])
            min_col = min(col_fill, key=lambda x: col_fill[x])

            # If difference is significant, try to move a student
            if col_fill[max_col] - col_fill[min_col] > 2:
                # Try to find a student in max_col that can be moved to min_col
                for student in rooms[r_name][max_col][:]:  # Iterate over copy
                    # Check if student can be placed in min_col
                    is_adj_safe, _ = check_adjacent_constraints(rooms[r_name], min_col, student)

                    if is_adj_safe:
                        # Move student
                        rooms[r_name][max_col].remove(student)
                        rooms[r_name][min_col].append(student)
                        moved = True
                        break

            if not moved:
                break

    # Calculate total placed
    total_placed = 0
    for r_name in rooms:
        for blk in BLOCK_ORDER:
            total_placed += len(rooms[r_name][blk])

    left = total_students - total_placed

    return rooms, left


# ===================== EXCEL READING =====================
def read_excel_file(file_path):
    """Read Excel file and return data"""
    try:
        print(f"Reading Excel file: {file_path}")

        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        col_map = {}
        for col in df.columns:
            col_lower = col.lower()
            if any(x in col_lower for x in ['register', 'reg', 'roll']):
                col_map['reg'] = col
            elif 'branch' in col_lower:
                col_map['branch'] = col
            elif any(x in col_lower for x in ['course', 'subject']):
                col_map['course'] = col

        print(f"Found columns: {col_map}")

        if 'reg' not in col_map or 'branch' not in col_map or 'course' not in col_map:
            print("❌ Missing required columns")
            return []

        raw_data = []

        for idx, row in df.iterrows():
            try:
                reg_no = str(row[col_map['reg']]).strip()
                branch_full = str(row[col_map['branch']]).strip()
                course_full = str(row[col_map['course']]).strip()

                if not reg_no or reg_no.upper() == 'NAN':
                    continue
                if not branch_full or branch_full.upper() == 'NAN':
                    continue
                if not course_full or course_full.upper() == 'NAN':
                    continue

                # Extract branch code
                branch_full_upper = branch_full.upper()
                if 'COMPUTER SCIENCE' in branch_full_upper:
                    branch = 'CSE'
                elif 'INFORMATION TECHNOLOGY' in branch_full_upper:
                    branch = 'IT'
                elif 'CIVIL' in branch_full_upper:
                    branch = 'CE'
                elif 'ELECTRONICS & COMMUNICATION' in branch_full_upper:
                    branch = 'EC'
                elif 'ELECTRONICS AND COMPUTER' in branch_full_upper:
                    branch = 'ECE'
                elif 'APPLIED ELECTRONICS' in branch_full_upper:
                    branch = 'AE'
                elif 'MECHANICAL' in branch_full_upper:
                    branch = 'ME'
                elif 'ELECTRICAL' in branch_full_upper:
                    branch = 'EE'
                elif 'CHEMICAL' in branch_full_upper:
                    branch = 'CH'
                elif 'BIOTECH' in branch_full_upper:
                    branch = 'BT'
                elif 'MATERIALS' in branch_full_upper:
                    branch = 'MT'
                else:
                    branch = branch_full[:3]

                # Keep the full course name with code for display
                course = course_full.strip()

                raw_data.append((reg_no, branch, course))

            except Exception as e:
                continue

        print(f"✅ Read {len(raw_data)} students from Excel")

        # Show statistics
        subject_counts = Counter([c for _, _, c in raw_data])
        print(f"\nTotal subjects: {len(subject_counts)}")
        print("Subject distribution:")
        for subject, count in subject_counts.most_common():
            print(f"  {subject}: {count}")

        return raw_data

    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return []


# ===================== REPORTING =====================
def save_report_to_files(rooms, original_counts, slot_date_folder, slot_file_path, total_actual_students):
    """Save seating arrangement report to files"""
    try:
        total_placed = 0
        arrangement = {
            'rooms': {},
            'summary': {
                'total_rooms': 0,
                'best_leftovers': 0,
                'best_difference': 0,
                'target_difference': 6,
                'actual_difference': 0,
                'average': 0,
                'min': 0,
                'max': 0
            },
            'student_count': total_actual_students,
            'subject_counts': original_counts,
            'elective_type': 'open_elective'
        }

        # Create folder if it doesn't exist
        os.makedirs(slot_date_folder, exist_ok=True)

        # Save as text report
        report_path = os.path.join(slot_date_folder, 'seating_report.txt')
        with open(report_path, "w", encoding="utf-8") as f:
            active_rooms = {k: v for k, v in rooms.items() if sum(len(col) for col in v.values()) > 0}
            room_names = sorted(active_rooms.keys(), key=lambda x: int(x.replace("Room", "")))

            arrangement['summary']['total_rooms'] = len(active_rooms)
            room_totals = []

            # First, build the room structure
            for r_name in room_names:
                r_data = active_rooms[r_name]
                occ = sum(len(col) for col in r_data.values())
                subjs = {st['subj'] for col in r_data.values() for st in col}
                depts = {st['branch'] for col in r_data.values() for st in col}
                room_totals.append(occ)
                total_placed += occ

                arrangement['rooms'][r_name] = {
                    'total': occ,
                    'blocks': {},
                    'subjects': list(subjs),
                    'departments': list(depts)
                }

                # Map columns to display block names
                for blk_orig in BLOCK_ORDER:
                    blk_mapped = BLOCK_MAPPING[blk_orig]
                    students = [st['roll'] for st in r_data[blk_orig]]
                    arrangement['rooms'][r_name]['blocks'][blk_mapped] = {
                        'students': students,
                        'capacity': COL_CAPACITY[blk_orig],
                        'count': len(students)
                    }

            # Update summary
            arrangement['summary']['best_leftovers'] = total_actual_students - total_placed

            # Calculate statistics
            if room_totals:
                avg = sum(room_totals) / len(room_totals)
                min_val = min(room_totals)
                max_val = max(room_totals)
                diff = max_val - min_val

                arrangement['summary']['average'] = avg
                arrangement['summary']['min'] = min_val
                arrangement['summary']['max'] = max_val
                arrangement['summary']['actual_difference'] = diff
                arrangement['summary']['best_difference'] = diff

            # Create QP summary
            room_wise = []
            subject_summary = defaultdict(int)

            for r_name in room_names:
                r_data = active_rooms[r_name]
                room_subjects = defaultdict(int)

                for blk_orig in BLOCK_ORDER:
                    for student in r_data[blk_orig]:
                        subject = student['full_subj']
                        room_subjects[subject] += 1
                        subject_summary[subject] += 1

                for subject, count in room_subjects.items():
                    room_wise.append({
                        'Room': r_name,
                        'Subject': subject,
                        'Student Count': count
                    })

            room_wise.sort(key=lambda x: (int(re.search(r'\d+', x['Room']).group()), x['Subject']))

            arrangement['qp_summary'] = {
                'room_wise': room_wise,
                'subject_summary': dict(subject_summary),
                'total_students': total_placed
            }

            # Write text report with column fill percentages
            for r_name in room_names:
                r_data = active_rooms[r_name]
                occ = sum(len(col) for col in r_data.values())
                subjs = {st['subj'] for col in r_data.values() for st in col}
                depts = {st['branch'] for col in r_data.values() for st in col}

                f.write(f"\n{'=' * 100}\n")
                f.write(
                    f"{r_name.upper()} | TOTAL: {occ}/{ROOM_MAX_CAP} | SUBJECTS: {len(subjs)} | DEPTS: {len(depts)}\n")
                f.write(f"{'=' * 100}\n")

                # Column headers with fill status
                display_headers = []
                for blk in BLOCK_ORDER:
                    fill = f"{len(r_data[blk])}/{COL_CAPACITY[blk]}"
                    display_headers.append(f"{BLOCK_MAPPING[blk]} ({fill})")

                f.write(f"{'Row':<4} | " + " | ".join([f"{h:<18}" for h in display_headers]) + "\n")
                f.write(f"{'-' * 100}\n")

                # Write rows
                max_rows = max(COL_CAPACITY.values())
                for i in range(max_rows):
                    row = f"{i + 1:<4} | "
                    for blk in BLOCK_ORDER:
                        if i < len(r_data[blk]):
                            st = r_data[blk][i]
                            cell = f"{st['roll']} ({st['branch']})"
                            row += f"{cell:<18} | "
                        elif i < COL_CAPACITY[blk]:
                            row += f"{'--':<18} | "
                        else:
                            row += f"{'':<18} | "
                    f.write(row + "\n")

                f.write(f"{'-' * 100}\n")

                # Subject and department counts
                subj_stats = Counter([st['subj'] for col in r_data.values() for st in col])
                dept_stats = Counter([st['branch'] for col in r_data.values() for st in col])

                f.write("Subjects in this room:\n")
                for s, c in subj_stats.items():
                    full_subj_name = next((st['full_subj'] for col in r_data.values() for st in col if st['subj'] == s),
                                          s)
                    f.write(f"  {full_subj_name}: {c} students\n")

                f.write("\nDepartments in this room:\n")
                for d, c in dept_stats.items():
                    f.write(f"  {d}: {c} students\n")

                # Column-wise subject distribution
                f.write("\nColumn-wise subject distribution:\n")
                for blk in BLOCK_ORDER:
                    col_subjs = set()
                    for st in r_data[blk]:
                        col_subjs.add(st['subj'])
                    f.write(f"  {BLOCK_MAPPING[blk]}: {', '.join(col_subjs) if col_subjs else 'Empty'}\n")
                f.write("\n")

            f.write(f"\n{'=' * 100}\n")
            f.write(f"SUMMARY\n")
            f.write(f"{'=' * 100}\n")
            f.write(f"Total students in file: {total_actual_students}\n")
            f.write(f"Total students placed: {total_placed}\n")
            f.write(f"Leftover students: {total_actual_students - total_placed}\n")
            f.write(f"Rooms used: {len(active_rooms)}\n")

            if room_totals:
                f.write(f"Capacity difference: {diff} (Target: ≤6)\n")
                f.write(f"Average per room: {avg:.1f}\n")
                f.write(f"Min room: {min_val}, Max room: {max_val}\n")

            # Check all constraints
            f.write(f"\n{'=' * 100}\n")
            f.write(f"CONSTRAINT VERIFICATION\n")
            f.write(f"{'=' * 100}\n")

            violations_subj = 0
            violations_dept = 0
            col_capacity_violations = 0
            max_subj_per_col_violations = 0

            for r_name in room_names:
                r_data = active_rooms[r_name]

                # Check column capacities
                for blk in BLOCK_ORDER:
                    if len(r_data[blk]) > COL_CAPACITY[blk]:
                        col_capacity_violations += 1
                        f.write(
                            f"⚠️ Capacity violation in {r_name}, {BLOCK_MAPPING[blk]}: {len(r_data[blk])} > {COL_CAPACITY[blk]}\n")

                # Check adjacent constraints
                for blk in BLOCK_ORDER:
                    col_idx = BLOCK_ORDER.index(blk)
                    if col_idx > 0:
                        left_blk = BLOCK_ORDER[col_idx - 1]
                        for st in r_data[blk]:
                            for left_st in r_data[left_blk]:
                                if st['subj'] == left_st['subj']:
                                    violations_subj += 1
                                    f.write(
                                        f"⚠️ Subject violation in {r_name}: {st['subj']} in {BLOCK_MAPPING[blk]} and {BLOCK_MAPPING[left_blk]}\n")
                                if st['branch'] == left_st['branch']:
                                    violations_dept += 1
                                    f.write(
                                        f"⚠️ Department violation in {r_name}: {st['branch']} in {BLOCK_MAPPING[blk]} and {BLOCK_MAPPING[left_blk]}\n")

                # Check max subjects per column
                for blk in BLOCK_ORDER:
                    col_subjs = {st['subj'] for st in r_data[blk]}
                    if len(col_subjs) > MAX_SUBJ_PER_COL:
                        max_subj_per_col_violations += 1
                        f.write(
                            f"⚠️ Max subjects violation in {r_name}, {BLOCK_MAPPING[blk]}: {len(col_subjs)} > {MAX_SUBJ_PER_COL}\n")

            if violations_subj == 0:
                f.write("✅ No adjacent same subjects found!\n")
            if violations_dept == 0:
                f.write("✅ No adjacent same departments found!\n")
            if col_capacity_violations == 0:
                f.write("✅ All column capacities respected!\n")
            if max_subj_per_col_violations == 0:
                f.write(f"✅ Max {MAX_SUBJ_PER_COL} subjects per column respected!\n")

        # Save JSON
        json_path = os.path.join(slot_date_folder, 'seating_arrangement.json')
        with open(json_path, 'w') as f:
            json.dump(arrangement, f, indent=2)

        # Save QP counts
        qp_path = os.path.join(slot_date_folder, 'qp_counts.txt')
        with open(qp_path, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("QUESTION PAPER COUNT REPORT\n")
            f.write("=" * 60 + "\n\n")

            f.write("ROOM-WISE DISTRIBUTION\n")
            f.write("-" * 60 + "\n")
            for item in arrangement['qp_summary']['room_wise']:
                f.write(f"{item['Room']}: {item['Subject']} - {item['Student Count']}\n")

            f.write("\n" + "=" * 60 + "\n")
            f.write("SUBJECT SUMMARY\n")
            f.write("=" * 60 + "\n")
            for subject, count in sorted(arrangement['qp_summary']['subject_summary'].items()):
                f.write(f"{subject}: {count}\n")

            f.write(f"\nTOTAL STUDENTS PLACED: {arrangement['qp_summary']['total_students']}\n")
            f.write("=" * 60 + "\n")

        return arrangement

    except Exception as e:
        print(f"❌ Error saving reports: {e}")
        import traceback
        traceback.print_exc()
        return None


# ===================== MAIN ALGORITHM =====================
def generate_open_elective_arrangement(slot_file_path, slot_date_folder):
    """Main function for Flask - Open Elective"""
    try:
        print(f"\n{'=' * 80}")
        print(f"OPEN ELECTIVE ALGORITHM")
        print(f"Target room difference: ≤6")
        print(f"Rules:")
        print(f"  - No same subject in adjacent columns")
        print(f"  - No same department in adjacent columns")
        print(f"  - Max {MAX_SUBJ_PER_COL} subjects per column")
        print(f"  - Uniform column filling")
        print(f"{'=' * 80}\n")

        # 1. Read Excel file
        raw_data = read_excel_file(slot_file_path)
        if not raw_data:
            return {'error': 'No valid student data found', 'elective_type': 'open_elective'}

        total_actual_students = len(raw_data)
        print(f"\n✅ Successfully read {total_actual_students} students")

        # 2. Process data
        subj_map, elective_counts, dept_map = extract_elective_counts(raw_data)
        total_students = sum(elective_counts.values())

        print(f"\n📊 Statistics:")
        print(f"  Total students: {total_students}")
        print(f"  Total subjects: {len(elective_counts)}")

        # Show all subjects
        print(f"\n📚 Subject distribution:")
        for subject, count in sorted(elective_counts.items(), key=lambda x: x[1], reverse=True):
            full_name = next((s[2] for s in raw_data if re.sub(r'\s*\([^)]*\)', '', s[2]).strip() == subject), subject)
            print(f"  {full_name}: {count}")

        # Show department distribution per subject
        print(f"\n🏢 Department distribution per subject:")
        for subject in sorted(elective_counts.keys()):
            if subject in dept_map:
                print(f"  {subject}: {', '.join(dept_map[subject])}")

        # 3. Calculate room count
        min_rooms_needed = math.ceil(total_students / ROOM_MAX_CAP)
        ideal_rooms = math.ceil(total_students / 30)  # Target ~30 per room

        print(f"\n🏢 Room calculation:")
        print(f"  Total students: {total_students}")
        print(f"  Max capacity per room: {ROOM_MAX_CAP}")
        print(f"  Minimum rooms needed: {min_rooms_needed}")
        print(f"  Ideal rooms for uniform distribution: {ideal_rooms}")

        best_allocation = None
        best_diff = float('inf')
        best_left = float('inf')
        best_rooms_used = 0
        best_col_balance = float('inf')  # Track column balance

        # 4. Try different room counts
        room_counts_to_try = list(range(max(min_rooms_needed, ideal_rooms - 2), min(20, ideal_rooms + 3)))
        print(f"  Trying room counts: {room_counts_to_try}")

        for rooms_try in room_counts_to_try:
            print(f"\n  Trying {rooms_try} rooms...")
            found_good_for_this_count = False

            for attempt in range(1000):  # More attempts for better distribution
                # Create a fresh deep copy for each attempt
                fresh_subj_map = copy.deepcopy(subj_map)

                final_rooms, left = generate_allocation(fresh_subj_map, elective_counts, dept_map, rooms_try)

                # Calculate room totals for active rooms
                room_totals = []
                col_balance_score = 0  # Track how evenly columns are filled

                for room_data in final_rooms.values():
                    occ = sum(len(col) for col in room_data.values())
                    if occ > 0:
                        room_totals.append(occ)

                        # Calculate column balance for this room
                        col_fills = [len(room_data[blk]) / COL_CAPACITY[blk] for blk in BLOCK_ORDER]
                        col_balance_score += max(col_fills) - min(col_fills)

                if not room_totals:
                    continue

                # Calculate difference
                diff = max(room_totals) - min(room_totals)
                avg = sum(room_totals) / len(room_totals)

                # Track best solution (prioritize zero leftovers, then small difference, then column balance)
                if left == 0:
                    if diff < best_diff or (diff == best_diff and col_balance_score < best_col_balance):
                        best_diff = diff
                        best_allocation = copy.deepcopy(final_rooms)
                        best_rooms_used = len(room_totals)
                        best_left = 0
                        best_col_balance = col_balance_score
                        print(
                            f"    ✓ Attempt {attempt + 1}: diff = {diff}, avg = {avg:.1f}, col balance = {col_balance_score:.2f}")

                        if diff <= 6:
                            print(f"      → Found good solution!")
                            found_good_for_this_count = True
                            # Don't break immediately, continue looking for even better column balance
                elif left < best_left:
                    best_left = left
                    best_diff = diff
                    best_allocation = copy.deepcopy(final_rooms)
                    best_rooms_used = len(room_totals)
                    best_col_balance = col_balance_score
                    print(f"    ⚠ Attempt {attempt + 1}: {left} students left, diff = {diff}, avg = {avg:.1f}")

                if (attempt + 1) % 200 == 0:
                    print(f"    ... {attempt + 1} attempts")

            if found_good_for_this_count:
                break

        # 5. Final result
        if best_allocation:
            # Calculate total placed students
            total_placed = 0
            for room_data in best_allocation.values():
                total_placed += sum(len(col) for col in room_data.values())

            rooms_used = len([r for r in best_allocation.values() if sum(len(col) for col in r.values()) > 0])

            print(f"\n{'=' * 80}")
            print(f"✅ SUCCESS!")
            print(f"  Total students in file: {total_actual_students}")
            print(f"  Students placed: {total_placed}")
            print(f"  Leftover: {total_actual_students - total_placed}")
            print(f"  Rooms used: {rooms_used}")
            print(f"  Capacity difference: {best_diff}")
            print(f"{'=' * 80}\n")

            # 6. Save results
            arrangement = save_report_to_files(best_allocation, elective_counts, slot_date_folder, slot_file_path,
                                               total_actual_students)

            if arrangement:
                return arrangement
            else:
                return {'error': 'Failed to save arrangement files', 'elective_type': 'open_elective'}
        else:
            print(f"\n❌ No valid allocation found")
            return {'error': 'No valid allocation found', 'elective_type': 'open_elective'}

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return {'error': str(e), 'elective_type': 'open_elective'}


# ===================== TEST =====================
if __name__ == "__main__":
    # Simple test
    print("Testing open elective algorithm...")

    test_data = [
        ("LBT22CS045", "CSE", "PYTHON FOR EVERYONE (OEL352)"),
        ("LBT22CS067", "CSE", "PYTHON FOR EVERYONE (OEL352)"),
        ("LBT22EC008", "EC", "PYTHON FOR EVERYONE (OEL352)"),
        ("LBT22EC012", "EC", "PYTHON FOR EVERYONE (OEL352)"),
        ("LBT22IT009", "IT", "PYTHON FOR EVERYONE (OEL352)"),
        ("LBT22ME005", "ME", "PYTHON FOR EVERYONE (OEL352)"),
    ]

    print(f"Test with {len(test_data)} students")

    subj_map, elective_counts, dept_map = extract_elective_counts(test_data)
    rooms_try = math.ceil(len(test_data) / 30)

    # Create a fresh copy for test
    fresh_subj_map = copy.deepcopy(subj_map)
    final_rooms, left = generate_allocation(fresh_subj_map, elective_counts, dept_map, rooms_try)

    if left == 0:
        print(f"✅ Test successful!")
    else:
        print(f"❌ Test failed: {left} students left")