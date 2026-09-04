import os
import re
import pandas as pd
import math
import random
import json
from collections import defaultdict, Counter
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, jsonify
from flask_session import Session
from werkzeug.utils import secure_filename
from datetime import datetime

# Import modules
from constraint_handler import ConstraintHandler
from rebalancer import Rebalancer
from series_allocation import SeriesAllocator

# Import the Excel converter
try:
    from txt_to_excel_converter import convert_txt_to_excel
    TXT_TO_EXCEL_AVAILABLE = True
except ImportError:
    print("WARNING: txt_to_excel_converter.py not found. Excel generation will be disabled.")
    TXT_TO_EXCEL_AVAILABLE = False

# Import PDF generator
try:
    from pdf_generator import create_seating_pdf, create_qp_summary_pdf, create_master_pdf
    PDF_AVAILABLE = True
except ImportError:
    print("WARNING: pdf_generator.py not found. PDF generation will be disabled.")
    PDF_AVAILABLE = False

# Import Series PDF generator
try:
    from series_pdf_generator import save_series_files
    SERIES_PDF_AVAILABLE = True
except ImportError:
    print("WARNING: series_pdf_generator.py not found. Series file generation will be disabled.")
    SERIES_PDF_AVAILABLE = False

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

# Configure server-side session
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
app.config['SESSION_FILE_DIR'] = './flask_session'
app.config['SESSION_FILE_THRESHOLD'] = 500
app.config['SESSION_FILE_MODE'] = 384

# Initialize Session
Session(app)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Create session directory if it doesn't exist
os.makedirs('./flask_session', exist_ok=True)
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Initialize modules
constraint_handler = ConstraintHandler()
rebalancer = Rebalancer(constraint_handler)
series_allocator = SeriesAllocator()

# Constants
BLOCK_ORDER = constraint_handler.BLOCK_ORDER
BLOCK_CAPACITY = constraint_handler.BLOCK_CAPACITY
MAX_TOTAL_ROOM = constraint_handler.MAX_TOTAL_ROOM
BRANCH_MAP = constraint_handler.BRANCH_MAP
CLASS_GROUP = constraint_handler.CLASS_GROUP
TARGET_MAX_DIFFERENCE = 4
MAX_ATTEMPTS = 1000
MAX_PERFECT_ATTEMPTS = 100

# Elective algorithms
PROGRAM_ELECTIVE_AVAILABLE = False
OPEN_ELECTIVE_AVAILABLE = False

try:
    from program_elect import generate_program_elective_arrangement
    PROGRAM_ELECTIVE_AVAILABLE = True
    print("✓ Program elective algorithm loaded successfully")
except ImportError:
    print("Warning: program_elect.py not found")
    PROGRAM_ELECTIVE_AVAILABLE = False

try:
    from open_elect import generate_open_elective_arrangement
    OPEN_ELECTIVE_AVAILABLE = True
    print("✓ Open elective algorithm loaded successfully")
except ImportError:
    print("Warning: open_elect.py not found")
    OPEN_ELECTIVE_AVAILABLE = False


# =============================
# 1. UPLOAD & SORTING FUNCTIONS
# =============================

def normalize_columns(df):
    df.columns = df.columns.astype(str).str.replace('"', '', regex=False).str.replace('\t', '', regex=False).str.strip()
    return df


def extract_student_info(df, student_col):
    def extract_name(student):
        match = re.match(r'(.+?)\(', str(student))
        return match.group(1).strip() if match else str(student)

    def extract_regno(student):
        match = re.search(r'\(([^)]+)\)', str(student))
        return match.group(1).strip() if match else ""

    df['Student Name'] = df[student_col].apply(extract_name)
    df['Register No'] = df[student_col].apply(extract_regno)
    return df


def extract_sorting_keys(df):
    def extract_year(reg):
        match = re.search(r'(?:LLBT|LBT)(\d{2})', str(reg))
        return int(match.group(1)) if match else 99

    def extract_serial(reg):
        match = re.search(r'([A-Z]{2})(\d{3})$', str(reg))
        return int(match.group(2)) if match else 999

    df['_year'] = df['Register No'].apply(extract_year)
    df['_serial'] = df['Register No'].apply(extract_serial)
    return df


def roll_key(r):
    return constraint_handler.roll_key(r)


def calculate_room_difference(rooms_data):
    totals = [sum(b.get("qty", 0) for b in room.values()) for room in rooms_data.values()]
    return max(totals) - min(totals) if totals else float('inf')


def create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path):
    """Create a simple Excel file as fallback if the main converter fails"""
    try:
        import pandas as pd

        with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
            summary_data = {
                'Metric': ['Total Students', 'Total Rooms', 'Average per Room', 'Min per Room', 'Max per Room',
                           'Room Difference'],
                'Value': [
                    arrangement.get('student_count', 0),
                    arrangement['summary'].get('total_rooms', 0),
                    round(arrangement['summary'].get('average', 0), 2),
                    arrangement['summary'].get('min', 0),
                    arrangement['summary'].get('max', 0),
                    arrangement['summary'].get('actual_difference', 0)
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

            if 'qp_summary' in arrangement and arrangement['qp_summary']:
                qp_data = []
                for item in arrangement['qp_summary'].get('room_wise', []):
                    qp_data.append({
                        'Room': item['Room'],
                        'Subject': item['Subject'],
                        'Count': item['Student Count']
                    })
                if qp_data:
                    pd.DataFrame(qp_data).to_excel(writer, sheet_name='QP Counts', index=False)

            for room_name, room_data in arrangement['rooms'].items():
                room_data_rows = []
                max_rows = max([len(block.get('students', [])) for block in room_data['blocks'].values()] or [0])

                for i in range(max_rows):
                    row = {'Row': i + 1}
                    for block_name in ['Left1', 'Left3', 'Middle2', 'Right1', 'Right3']:
                        block = room_data['blocks'].get(block_name, {})
                        students = block.get('students', [])
                        row[block_name] = students[i] if i < len(students) else '--'
                    room_data_rows.append(row)

                if room_data_rows:
                    sheet_name = room_name[:31]
                    pd.DataFrame(room_data_rows).to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✓ Fallback Excel file created: {excel_output_path}")
        return True
    except Exception as e:
        print(f"✗ Fallback Excel creation failed: {e}")
        return False


def get_room_to_hall_map():
    """Return the room-to-hall number mapping"""
    return {
        "Room1": "108", "Room2": "109", "Room3": "111", "Room4": "112",
        "Room5": "208", "Room6": "209", "Room7": "211", "Room8": "212",
        "Room9": "308", "Room10": "309", "Room11": "311", "Room12": "312",
    }


def save_arrangement_files(arrangement, slot_date_folder, slot_file_path):
    """Save arrangement files (JSON, TXT report, QP counts, Master PDF)"""
    try:
        # Save arrangement as JSON
        arrangement_json_path = os.path.join(slot_date_folder, 'seating_arrangement.json')
        with open(arrangement_json_path, 'w') as f:
            json.dump(arrangement, f, indent=2)

        # Save as text report (for printing)
        report_path = os.path.join(slot_date_folder, 'seating_report.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("AUTOMATED SEATING ARRANGEMENT REPORT\n")
            if 'elective_type' in arrangement:
                f.write(f"Elective Type: {arrangement['elective_type'].replace('_', ' ').title()}\n")
            f.write("=" * 80 + "\n\n")

            if 'qp_summary' in arrangement:
                f.write("QUESTION PAPER COUNT SUMMARY\n")
                f.write("-" * 40 + "\n")
                for subject, count in arrangement['qp_summary']['subject_summary'].items():
                    f.write(f"{subject}: {count} students\n")
                f.write(f"Total Students: {arrangement['qp_summary']['total_students']}\n")
                f.write("\n" + "=" * 80 + "\n\n")

            for room_name, room_data in arrangement['rooms'].items():
                f.write("=" * 80 + "\n")
                f.write(f" {room_name} | TOTAL: {room_data['total']} | BLOCK CAPACITY: 7-7-6-7-7 \n")
                f.write("=" * 80 + "\n")
                f.write(
                    "Row   | Left1 (7)       | Left3 (7)       | Middle2 (6)     | Right1 (7)      | Right3 (7)     \n")
                f.write("-" * 95 + "\n")

                for row in range(1, 8):
                    left1 = room_data['blocks']['Left1']['students'][row - 1] if row <= len(
                        room_data['blocks']['Left1']['students']) else '--'
                    left3 = room_data['blocks']['Left3']['students'][row - 1] if row <= len(
                        room_data['blocks']['Left3']['students']) else '--'
                    middle2 = room_data['blocks']['Middle2']['students'][row - 1] if row <= 6 and row <= len(
                        room_data['blocks']['Middle2']['students']) else '--'
                    right1 = room_data['blocks']['Right1']['students'][row - 1] if row <= len(
                        room_data['blocks']['Right1']['students']) else '--'
                    right3 = room_data['blocks']['Right3']['students'][row - 1] if row <= len(
                        room_data['blocks']['Right3']['students']) else '--'

                    left1 = str(left1).ljust(15)
                    left3 = str(left3).ljust(15)
                    middle2 = str(middle2).ljust(15)
                    right1 = str(right1).ljust(15)
                    right3 = str(right3).ljust(15)

                    if row == 7:
                        f.write(f"{row:<5} | {left1} | {left3} | {' '.ljust(15)} | {right1} | {right3} |\n")
                    else:
                        f.write(f"{row:<5} | {left1} | {left3} | {middle2} | {right1} | {right3} |\n")

                f.write("-" * 95 + "\n")
                f.write(
                    f"Block Usage: Left1: {room_data['blocks']['Left1']['count']}/7, Left3: {room_data['blocks']['Left3']['count']}/7, Middle2: {room_data['blocks']['Middle2']['count']}/6, Right1: {room_data['blocks']['Right1']['count']}/7, Right3: {room_data['blocks']['Right3']['count']}/7\n")
                f.write(f"Subjects: {', '.join(room_data['subjects'])}\n\n")

            f.write("=" * 80 + "\n")
            f.write("SUMMARY\n")
            f.write("=" * 80 + "\n")
            f.write(f"Total Rooms: {arrangement['summary']['total_rooms']}\n")
            f.write(f"Total Students: {arrangement['student_count']}\n")
            f.write(f"Average per Room: {arrangement['summary']['average']:.1f}\n")
            f.write(f"Leftover Students: {arrangement['summary']['best_leftovers']}\n")
            f.write(f"Max Room Difference: {arrangement['summary']['actual_difference']}\n")

        if 'qp_summary' in arrangement:
            qp_path = os.path.join(slot_date_folder, 'qp_counts.txt')
            with open(qp_path, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("QUESTION PAPER COUNT REPORT\n")
                if 'elective_type' in arrangement:
                    f.write(f"Elective Type: {arrangement['elective_type'].replace('_', ' ').title()}\n")
                f.write("=" * 60 + "\n\n")
                f.write("ROOM-WISE DISTRIBUTION\n")
                f.write("-" * 60 + "\n")
                for item in arrangement['qp_summary']['room_wise']:
                    f.write(f"{item['Room']}: {item['Subject']} - {item['Student Count']} students\n")
                f.write("\n" + "=" * 60 + "\n")
                f.write("SUBJECT SUMMARY\n")
                f.write("=" * 60 + "\n")
                for subject, count in sorted(arrangement['qp_summary']['subject_summary'].items()):
                    f.write(f"{subject}: {count} students\n")
                f.write(f"\nTotal Students: {arrangement['qp_summary']['total_students']}\n")
                f.write("=" * 60 + "\n")

        if PDF_AVAILABLE:
            try:
                from txt_to_excel_converter import parse_txt, extract_metadata_from_txt
                with open(report_path, 'r', encoding='utf-8', errors='replace') as _f:
                    _txt = _f.read()
                _metadata = extract_metadata_from_txt(_txt)
                _meta, _rooms = parse_txt(_txt, _metadata)
                _hall_map = get_room_to_hall_map()
                master_pdf_path = os.path.join(slot_date_folder, 'master_list.pdf')
                _ok = create_master_pdf(
                    _rooms, _meta, _hall_map, master_pdf_path,
                    institution=_metadata.get(
                        'institution',
                        'LBS INSTITUTE OF TECHNOLOGY FOR WOMEN, Poojappura, Thiruvananthapuram'),
                    exam_title=_metadata.get('exam_title', '')
                )
                if _ok:
                    print(f"✓ Master PDF generated: {master_pdf_path}")
            except Exception as _e:
                print(f"Master PDF generation error: {_e}")

        print(f"✓ Arrangement files saved in {slot_date_folder}")
        return True
    except Exception as e:
        print(f"Error saving arrangement files: {e}")
        return False


# =============================
# 2. PROCESS MASTER FILE
# =============================

def process_master_file(file_path, exam_type, semester, month_year, elective_type='general'):
    """Process master file and create sorted files with proper folder structure"""
    try:
        df = pd.read_excel(file_path)
        df = normalize_columns(df)

        student_col = None
        for col in df.columns:
            if 'student' in col.lower() or 'name' in col.lower():
                student_col = col
                break

        if not student_col:
            raise ValueError("Could not find student column")

        branch_col = next((col for col in df.columns if 'branch' in col.lower()), 'Branch Name')
        slot_col = next((col for col in df.columns if 'slot' in col.lower()), 'Slot')
        course_col = next((col for col in df.columns if 'course' in col.lower()), 'Course')
        exam_date_col = next((col for col in df.columns if 'exam' in col.lower() and 'date' in col.lower()),
                             'Exam Date')

        df = extract_student_info(df, student_col)
        df = extract_sorting_keys(df)

        final_df = df[[
            'Student Name',
            'Register No',
            branch_col,
            slot_col,
            course_col,
            exam_date_col if exam_date_col in df.columns else 'Exam Date'
        ]].copy()

        final_df.columns = [
            'Student',
            'Register No',
            'Branch Name',
            'Slot',
            'Course',
            'Exam Date'
        ]

        final_df['_year'] = df['_year']
        final_df['_serial'] = df['_serial']

        final_df = final_df.sort_values(
            by=['Branch Name', 'Slot', '_year', '_serial'],
            ascending=[True, True, True, True],
            kind='mergesort'
        )

        final_df.insert(0, 'Sl.No', range(1, len(final_df) + 1))
        final_df = final_df.drop(columns=['_year', '_serial'])

        if elective_type != 'general':
            base_path = os.path.join(
                app.config['UPLOAD_FOLDER'],
                str(exam_type),
                str(elective_type),
                str(semester),
                str(month_year)
            )
        else:
            base_path = os.path.join(
                app.config['UPLOAD_FOLDER'],
                str(exam_type),
                str(semester),
                str(month_year)
            )

        master_dir = os.path.join(base_path, 'master_list')
        os.makedirs(master_dir, exist_ok=True)

        master_file_path = os.path.join(master_dir, 'Master_Sorted_List.xlsx')
        final_df.to_excel(master_file_path, index=False)

        slot_files = {}

        for slot in final_df['Slot'].dropna().unique():
            slot_str = str(slot).strip()
            slot_folder = os.path.join(base_path, f'slot_{slot_str}')
            os.makedirs(slot_folder, exist_ok=True)

            slot_dates = final_df[final_df['Slot'] == slot_str]['Exam Date'].unique()

            for exam_date in slot_dates:
                if pd.isna(exam_date):
                    date_str = 'NoDate'
                else:
                    try:
                        date_str = pd.to_datetime(exam_date).strftime('%Y-%m-%d')
                    except:
                        date_str = str(exam_date)

                date_folder = os.path.join(slot_folder, date_str)
                os.makedirs(date_folder, exist_ok=True)

                slot_df = final_df[(final_df['Slot'] == slot_str) & (final_df['Exam Date'] == exam_date)].copy()
                if len(slot_df) > 0:
                    slot_df['Sl.No'] = range(1, len(slot_df) + 1)

                    slot_filename = f'Slot_{slot_str}_{date_str}_Sorted_List.xlsx'
                    slot_file_path = os.path.join(date_folder, slot_filename)
                    slot_df.to_excel(slot_file_path, index=False)

                    if slot_str not in slot_files:
                        slot_files[slot_str] = []
                    slot_files[slot_str].append({
                        'date': date_str,
                        'file_path': slot_file_path,
                        'folder': date_folder,
                        'student_count': len(slot_df)
                    })

        return {
            'success': True,
            'master_file': master_file_path,
            'slot_files': slot_files,
            'slots': list(final_df['Slot'].dropna().unique()),
            'elective_type': elective_type,
            'base_path': base_path
        }

    except Exception as e:
        print(f"Error in process_master_file: {e}")
        import traceback
        traceback.print_exc()
        return {
            'success': False,
            'error': str(e)
        }


# =============================
# 3. SEATING ALLOCATION FUNCTIONS
# =============================

def generate_allocation(classes, supply_data, subjects):
    pool = []
    for cls, count in classes.items():
        s_count = sum(supply_data.get(cls, {}).values())
        pool.extend([cls] * (count + s_count))

    random.shuffle(pool)
    current_pool = pool.copy()

    # ── FIX: single-subject rooms only use 20 seats (Middle2+Left1+Right3)
    #         so calculate num_rooms based on 20, not 31 ──────────────────
    unique_classes = set(pool)
    avg_per_room = 20 if len(unique_classes) == 1 else 34
    # ─────────────────────────────────────────────────────────────────────

    num_rooms = math.ceil(len(pool) / avg_per_room)
    rooms = {}

    target_per_room = []
    total_students = len(pool)
    for i in range(num_rooms):
        if i == num_rooms - 1:
            target_per_room.append(total_students - sum(target_per_room))
        else:
            base = total_students // num_rooms
            variation = random.randint(-3, 3)
            target_per_room.append(min(base + variation, MAX_TOTAL_ROOM))

    for r in range(num_rooms):
        if not current_pool:
            break

        room = f"Room{r + 1}"
        room_target = target_per_room[r]

        counts = Counter(current_pool)
        if not counts:
            break

        main_sub = max(counts.items(), key=lambda x: x[1])[0]
        main_group = CLASS_GROUP.get(main_sub)

        possible_secondary = []
        for cls, cnt in counts.items():
            if cls == main_sub:
                continue
            if CLASS_GROUP.get(cls) != main_group:
                possible_secondary.append((cls, cnt))

        possible_secondary.sort(key=lambda x: x[1], reverse=True)

        if possible_secondary:
            sec_sub = possible_secondary[0][0]
        else:
            sec_sub = main_sub

        main_count = counts[main_sub]
        sec_count = counts.get(sec_sub, 0)

        target_main = min(main_count, max(10, int(room_target * 0.65)))
        target_sec = min(sec_count, room_target - target_main)

        remaining = room_target - (target_main + target_sec)
        while remaining > 0:
            if target_main < main_count and target_main < 24:
                target_main += 1
            elif target_sec < sec_count and target_sec < 16:
                target_sec += 1
            else:
                break
            remaining = room_target - (target_main + target_sec)

        rooms[room] = {
            "sub_a": {"cls": main_sub, "qty": target_main},
            "sub_b": {"cls": sec_sub, "qty": target_sec}
        }

        for _ in range(target_main):
            if main_sub in current_pool:
                current_pool.remove(main_sub)

        if main_sub != sec_sub:
            for _ in range(target_sec):
                if sec_sub in current_pool:
                    current_pool.remove(sec_sub)

    return rooms, current_pool


def create_block_layout(rooms_data, subjects):
    final = {}

    for room_name, data in rooms_data.items():
        blocks = {b: {} for b in BLOCK_ORDER}

        a_cls = data["sub_a"]["cls"]
        a_qty = data["sub_a"]["qty"]

        if a_qty > 0:
            total_to_distribute = a_qty
            left1_cap = BLOCK_CAPACITY["Left1"]
            middle2_cap = BLOCK_CAPACITY["Middle2"]
            right3_cap = BLOCK_CAPACITY["Right3"]

            if total_to_distribute > 0:
                middle_qty = min(total_to_distribute, middle2_cap)
                if middle_qty > 0:
                    blocks["Middle2"] = {"cls": a_cls, "qty": middle_qty, "subject": subjects.get(a_cls)}
                    total_to_distribute -= middle_qty

            if total_to_distribute > 0:
                if total_to_distribute <= (left1_cap + right3_cap):
                    left_qty = min(math.ceil(total_to_distribute / 2), left1_cap)
                    right_qty = total_to_distribute - left_qty

                    if right_qty > right3_cap:
                        right_qty = right3_cap
                        left_qty = total_to_distribute - right_qty

                    if left_qty > 0:
                        blocks["Left1"] = {"cls": a_cls, "qty": left_qty, "subject": subjects.get(a_cls)}
                    if right_qty > 0:
                        blocks["Right3"] = {"cls": a_cls, "qty": right_qty, "subject": subjects.get(a_cls)}
                else:
                    left_qty = min(total_to_distribute, left1_cap)
                    if left_qty > 0:
                        blocks["Left1"] = {"cls": a_cls, "qty": left_qty, "subject": subjects.get(a_cls)}
                        total_to_distribute -= left_qty

                    if total_to_distribute > 0:
                        right_qty = min(total_to_distribute, right3_cap)
                        if right_qty > 0:
                            blocks["Right3"] = {"cls": a_cls, "qty": right_qty, "subject": subjects.get(a_cls)}

        b_cls = data["sub_b"]["cls"]
        b_qty = data["sub_b"]["qty"]

        if b_qty > 0 and b_cls != a_cls:
            total_to_distribute = b_qty
            left3_cap = BLOCK_CAPACITY["Left3"]
            right1_cap = BLOCK_CAPACITY["Right1"]

            if total_to_distribute <= (left3_cap + right1_cap):
                left3_qty = min(math.ceil(total_to_distribute / 2), left3_cap)
                right1_qty = total_to_distribute - left3_qty

                if right1_qty > right1_cap:
                    right1_qty = right1_cap
                    left3_qty = total_to_distribute - right1_qty

                if left3_qty > 0:
                    blocks["Left3"] = {"cls": b_cls, "qty": left3_qty, "subject": subjects.get(b_cls)}
                if right1_qty > 0:
                    blocks["Right1"] = {"cls": b_cls, "qty": right1_qty, "subject": subjects.get(b_cls)}
            else:
                left3_qty = min(total_to_distribute, left3_cap)
                if left3_qty > 0:
                    blocks["Left3"] = {"cls": b_cls, "qty": left3_qty, "subject": subjects.get(b_cls)}
                    total_to_distribute -= left3_qty

                if total_to_distribute > 0:
                    right1_qty = min(total_to_distribute, right1_cap)
                    if right1_qty > 0:
                        blocks["Right1"] = {"cls": b_cls, "qty": right1_qty, "subject": subjects.get(b_cls)}

        final[room_name] = blocks

    return final


def cleanup_leftovers(rooms, leftovers, subjects):
    if not leftovers:
        return rooms

    counts = Counter(leftovers)

    for cls, count in counts.items():
        room_order = sorted(rooms.keys(),
                            key=lambda x: sum(b.get("qty", 0) for b in rooms[x].values()))

        for room_name in room_order:
            if count <= 0:
                break

            blocks = rooms[room_name]
            current_total = sum(b.get("qty", 0) for b in blocks.values())

            if current_total >= MAX_TOTAL_ROOM:
                continue

            existing_blocks = []
            for block_name, block_data in blocks.items():
                if block_data and block_data.get("cls") == cls:
                    existing_blocks.append((block_name, block_data))

            if existing_blocks:
                for block_name, block_data in existing_blocks:
                    if count <= 0:
                        break
                    if block_data.get("qty", 0) < BLOCK_CAPACITY[block_name]:
                        add_qty = min(count, BLOCK_CAPACITY[block_name] - block_data["qty"])
                        block_data["qty"] += add_qty
                        count -= add_qty
            else:
                for block_name in BLOCK_ORDER:
                    if count <= 0:
                        break
                    if not blocks[block_name]:
                        subjects_in_room = len({b["cls"] for b in blocks.values() if b})
                        if subjects_in_room < 2:
                            add_qty = min(count, BLOCK_CAPACITY[block_name], MAX_TOTAL_ROOM - current_total)
                            blocks[block_name] = {
                                "cls": cls,
                                "qty": add_qty,
                                "subject": subjects.get(cls)
                            }
                            count -= add_qty
                            break

    return rooms


def generate_qp_counts(arrangement, slot_file_path):
    """Generate accurate QP counts by analyzing actual student data"""
    try:
        df = pd.read_excel(slot_file_path)

        branch_subject_map = {}
        for _, row in df.iterrows():
            reg_no = str(row.get("Register No", "")).strip()
            branch = str(row.get("Branch Name", "")).strip()
            course = str(row.get("Course", "")).strip()

            if branch and course:
                bcode = BRANCH_MAP.get(branch)
                if bcode and bcode not in branch_subject_map:
                    branch_subject_map[bcode] = course

        if 'branch_subject_map' in arrangement:
            for bcode, subject in arrangement['branch_subject_map'].items():
                if bcode not in branch_subject_map:
                    branch_subject_map[bcode] = subject

        room_student_map = {}
        for room_name, room_data in arrangement['rooms'].items():
            room_students = []
            for block_name, block_data in room_data['blocks'].items():
                room_students.extend(block_data['students'])
            room_student_map[room_name] = room_students

        qp_data = []
        subject_totals = {}

        branch_patterns = {
            'S7CSE': ['CS', 'CSE'],
            'S7IT': ['IT'],
            'S7EC': ['EC'],
            'S7ER': ['ECE', 'ER'],
            'S7CE': ['CE']
        }

        for room_name, students in room_student_map.items():
            room_subjects = {}

            for student_roll in students:
                if student_roll == '--' or not student_roll:
                    continue

                student_roll_upper = str(student_roll).upper()
                matched_branch = None

                for bcode, patterns in branch_patterns.items():
                    for pattern in patterns:
                        if pattern in student_roll_upper:
                            matched_branch = bcode
                            break
                    if matched_branch:
                        break

                if not matched_branch:
                    if 'LLBT' in student_roll_upper or 'LBT' in student_roll_upper:
                        match = re.search(r'(?:LLBT|LBT)\d{2}([A-Z]{2,3})', student_roll_upper)
                        if match:
                            dept_code = match.group(1)
                            if dept_code in ['CS', 'CSE']:
                                matched_branch = 'S7CSE'
                            elif dept_code == 'IT':
                                matched_branch = 'S7IT'
                            elif dept_code == 'EC':
                                matched_branch = 'S7EC'
                            elif dept_code in ['ECE', 'ER']:
                                matched_branch = 'S7ER'
                            elif dept_code == 'CE':
                                matched_branch = 'S7CE'

                if matched_branch and matched_branch in branch_subject_map:
                    subject_name = branch_subject_map[matched_branch]

                    if subject_name not in room_subjects:
                        room_subjects[subject_name] = 0
                    room_subjects[subject_name] += 1

                    if subject_name not in subject_totals:
                        subject_totals[subject_name] = 0
                    subject_totals[subject_name] += 1

            for subject_name, count in room_subjects.items():
                qp_data.append({
                    'Room': room_name,
                    'Subject': subject_name,
                    'Student Count': count
                })

        qp_data.sort(key=lambda x: (int(re.search(r'\d+', x['Room']).group()), x['Subject']))

        qp_summary = {
            'room_wise': qp_data,
            'subject_summary': subject_totals,
            'total_students': sum(subject_totals.values()) if subject_totals else 0
        }

        return qp_summary

    except Exception as e:
        print(f"Error in generate_qp_counts: {e}")
        return None


def generate_seating_arrangement(slot_file_path, slot_date_folder, elective_type='general'):
    """Generate seating arrangement based on elective type"""
    try:
        print(f"\n{'=' * 60}")
        print(f"Generating seating arrangement for: {slot_file_path}")
        print(f"Elective Type: {elective_type}")
        print(f"Date Folder: {slot_date_folder}")
        print(f"{'=' * 60}\n")

        # Extract the date from the folder name (YYYY-MM-DD format)
        folder_name = os.path.basename(slot_date_folder)
        exam_date_value = folder_name  # This is already in YYYY-MM-DD format
        print(f"Exam date from folder: {exam_date_value}")

        if elective_type == 'program_elective' and PROGRAM_ELECTIVE_AVAILABLE:
            print("Using program elective algorithm")
            arrangement = generate_program_elective_arrangement(slot_file_path, slot_date_folder)
        elif elective_type == 'open_elective' and OPEN_ELECTIVE_AVAILABLE:
            print("Using open elective algorithm")
            arrangement = generate_open_elective_arrangement(slot_file_path, slot_date_folder)
        else:
            print("Using general algorithm with CHM & RBM")

            df = pd.read_excel(slot_file_path)
            df.columns = df.columns.str.strip()

            student_data = constraint_handler.process_student_data(df, elective_type)

            MASTER_ROLLS = student_data['rolls']
            MASTER_SUBJECTS = student_data['subjects']
            classes_count = student_data['classes_count']
            supply_data = student_data['supply_data']
            total_students = student_data['total_students']

            best_solution_phase1 = None
            best_leftovers_phase1 = float('inf')
            best_difference_phase1 = float('inf')
            phase1_completed = False

            for attempt in range(MAX_ATTEMPTS):
                rooms_data, leftovers = generate_allocation(classes_count, supply_data, MASTER_SUBJECTS)
                block_layout = create_block_layout(rooms_data, MASTER_SUBJECTS)

                block_layout, rebalance_stats = rebalancer.rebalance(block_layout, elective_type)
                block_layout = rebalancer.cleanup_empty_blocks(block_layout)

                block_layout = cleanup_leftovers(block_layout, leftovers, MASTER_SUBJECTS)

                total_students = sum(len(rolls) for rolls in MASTER_ROLLS.values())
                allocated_students = sum(
                    sum(b.get("qty", 0) for b in room.values())
                    for room in block_layout.values()
                )
                current_leftovers = total_students - allocated_students
                current_difference = calculate_room_difference(block_layout)

                if current_leftovers < best_leftovers_phase1 or (
                        current_leftovers == best_leftovers_phase1 and current_difference < best_difference_phase1):
                    best_leftovers_phase1 = current_leftovers
                    best_difference_phase1 = current_difference
                    best_solution_phase1 = block_layout.copy()

                if current_leftovers == 0:
                    phase1_completed = True
                    break

            if not phase1_completed:
                best_solution = best_solution_phase1
                best_leftovers = best_leftovers_phase1
                best_difference = best_difference_phase1
            else:
                best_solution = best_solution_phase1
                best_leftovers = 0
                best_difference = best_difference_phase1

                for attempt in range(MAX_PERFECT_ATTEMPTS):
                    rooms_data, leftovers = generate_allocation(classes_count, supply_data, MASTER_SUBJECTS)
                    block_layout = create_block_layout(rooms_data, MASTER_SUBJECTS)

                    block_layout, rebalance_stats = rebalancer.rebalance(block_layout, elective_type)
                    block_layout = rebalancer.cleanup_empty_blocks(block_layout)

                    block_layout = cleanup_leftovers(block_layout, leftovers, MASTER_SUBJECTS)

                    total_students = sum(len(rolls) for rolls in MASTER_ROLLS.values())
                    allocated_students = sum(
                        sum(b.get("qty", 0) for b in room.values())
                        for room in block_layout.values()
                    )
                    current_leftovers = total_students - allocated_students
                    current_difference = calculate_room_difference(block_layout)

                    if current_leftovers == 0:
                        if current_difference < best_difference:
                            best_difference = current_difference
                            best_solution = block_layout.copy()

                        if current_difference <= TARGET_MAX_DIFFERENCE:
                            break

            rooms_data = best_solution
            working_rolls = {k: list(v) for k, v in MASTER_ROLLS.items()}

            arrangement = {
                'rooms': {},
                'summary': {
                    'total_rooms': len(rooms_data),
                    'best_leftovers': best_leftovers,
                    'best_difference': best_difference,
                    'target_difference': TARGET_MAX_DIFFERENCE
                },
                'student_count': sum(len(rolls) for rolls in MASTER_ROLLS.values()),
                'branch_subject_map': MASTER_SUBJECTS,
                'elective_type': elective_type
            }

            for r in sorted(rooms_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())):
                blocks = rooms_data[r]

                col_data = {b: [] for b in BLOCK_ORDER}
                for blk in BLOCK_ORDER:
                    b = blocks.get(blk)
                    if b and b.get("cls"):
                        for _ in range(b["qty"]):
                            if working_rolls[b["cls"]]:
                                col_data[blk].append(working_rolls[b["cls"]].pop(0))

                room_total = sum(len(v) for v in col_data.values())

                subjects_set = set()
                for b in blocks.values():
                    if b and b.get('cls'):
                        subjects_set.add(f"{b['cls']}: {b['subject']}")

                arrangement['rooms'][r] = {
                    'total': room_total,
                    'blocks': {},
                    'subjects': list(subjects_set)
                }

                for blk in BLOCK_ORDER:
                    arrangement['rooms'][r]['blocks'][blk] = {
                        'students': col_data[blk],
                        'capacity': BLOCK_CAPACITY[blk],
                        'count': len(col_data[blk])
                    }

            room_totals = []
            for r in arrangement['rooms']:
                room_totals.append(arrangement['rooms'][r]['total'])

            if room_totals:
                arrangement['summary']['average'] = sum(room_totals) / len(room_totals)
                arrangement['summary']['min'] = min(room_totals)
                arrangement['summary']['max'] = max(room_totals)
                arrangement['summary']['actual_difference'] = max(room_totals) - min(room_totals)

        if elective_type == 'general':
            qp_summary = generate_qp_counts(arrangement, slot_file_path)
            if qp_summary:
                arrangement['qp_summary'] = qp_summary

        save_arrangement_files(arrangement, slot_date_folder, slot_file_path)

        excel_output_path = os.path.join(slot_date_folder, 'seating_output.xlsx')
        if TXT_TO_EXCEL_AVAILABLE:
            try:
                txt_report_path = os.path.join(slot_date_folder, 'seating_report.txt')
                # Pass both folder path and exam date
                excel_result = convert_txt_to_excel(txt_report_path, excel_output_path,
                                                    folder_path=slot_date_folder,
                                                    exam_date=exam_date_value)
                if excel_result and excel_result.get('success'):
                    arrangement['excel_report'] = excel_result['excel_path']
                    arrangement['excel_metadata'] = excel_result.get('metadata', {})
                    print(f"✓ Excel file generated: {excel_result['excel_path']}")
                    print(f"  Date extracted: {excel_result['metadata'].get('date_display', 'Not found')}")
                else:
                    print(f"Excel conversion failed: {excel_result.get('error', 'Unknown error')}")
                    if create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path):
                        arrangement['excel_report'] = excel_output_path
            except Exception as e:
                print(f"Excel generation error: {e}")
                import traceback
                traceback.print_exc()
                if create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path):
                    arrangement['excel_report'] = excel_output_path
        else:
            if create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path):
                arrangement['excel_report'] = excel_output_path

        if PDF_AVAILABLE:
            try:
                # Get the date from the folder name (YYYY-MM-DD format) and convert to DD.MM.YYYY
                folder_name = os.path.basename(slot_date_folder)
                date_match = re.match(r'(\d{4})-(\d{2})-(\d{2})', folder_name)
                if date_match:
                    year = date_match.group(1)
                    month = date_match.group(2)
                    day = date_match.group(3)
                    exam_date_display = f"{day}.{month}.{year}"
                else:
                    exam_date_display = folder_name.replace('-', '.')

                pdf_path = os.path.join(slot_date_folder, 'seating_arrangement.pdf')
                success = create_seating_pdf(
                    arrangement['rooms'],
                    arrangement['summary'],
                    pdf_path,
                    exam_date=exam_date_display,
                    semester=elective_type.capitalize() if elective_type != 'general' else 'Regular'
                )
                if success:
                    arrangement['pdf_report'] = pdf_path
                    print(f"✓ PDF report generated: {pdf_path}")

                if 'qp_summary' in arrangement and arrangement['qp_summary']:
                    qp_pdf_path = os.path.join(slot_date_folder, 'qp_summary.pdf')
                    qp_success = create_qp_summary_pdf(
                        arrangement['qp_summary'],
                        qp_pdf_path,
                        exam_date=exam_date_display
                    )
                    if qp_success:
                        arrangement['qp_pdf'] = qp_pdf_path
                        print(f"✓ QP Summary PDF generated: {qp_pdf_path}")
            except Exception as e:
                print(f"PDF generation error: {e}")
                import traceback
                traceback.print_exc()

        return arrangement

    except Exception as e:
        print(f"Error in generate_seating_arrangement: {e}")
        import traceback
        traceback.print_exc()
        return {'error': str(e), 'elective_type': elective_type}

# =============================
# 4. SERIES EXAM FUNCTIONS
# =============================

def generate_series_arrangement(series_data, algorithm='general'):
    """Generate seating arrangement for series exam using SeriesAllocator"""
    try:
        print("\n--- generate_series_arrangement called ---")
        print(f"Series data type: {type(series_data)}")
        print(f"Series data keys: {series_data.keys() if isinstance(series_data, dict) else 'Not a dict'}")

        custom_rooms = series_data.get('rooms', None)
        print(f"Custom rooms (total capacities): {custom_rooms}")

        allocator = SeriesAllocator(custom_rooms=custom_rooms)

        print("Calling allocator.generate_series_arrangement...")
        arrangement = allocator.generate_series_arrangement(series_data)
        print(f"Arrangement generated: {type(arrangement)}")

        if isinstance(arrangement, dict):
            print(f"Arrangement keys: {arrangement.keys()}")
            if 'rooms' in arrangement:
                print(f"Number of rooms: {len(arrangement['rooms'])}")

        # ── Save all three output files into uploads/Series/<date>/ ──────────
        if isinstance(arrangement, dict) and 'rooms' in arrangement and SERIES_PDF_AVAILABLE:
            try:
                date_str = datetime.now().strftime('%Y-%m-%d')
                series_folder = os.path.join(
                    app.config['UPLOAD_FOLDER'], 'Series', date_str
                )
                exam_info = series_data.get('exam_info', None)
                saved = save_series_files(arrangement, series_folder, exam_info=exam_info)
                arrangement['output_files'] = saved
                arrangement['output_folder'] = series_folder
                print(f"✓ Series files saved to: {series_folder}")
                for key, path in saved.items():
                    print(f"   {key}: {path}")
            except Exception as file_err:
                import traceback
                print(f"Warning: could not save series output files: {file_err}")
                traceback.print_exc()
        # ─────────────────────────────────────────────────────────────────────

        return arrangement

    except Exception as e:
        print(f"Error in generate_series_arrangement: {e}")
        import traceback
        traceback.print_exc()
        return {'error': str(e)}


# =============================
# 5. FLASK ROUTES
# =============================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/set_exam_type', methods=['POST'])
def set_exam_type():
    exam_type = request.form.get('exam_type')
    if exam_type:
        session['exam_type'] = exam_type
        session.pop('series_data', None)
        session.pop('series_arrangement', None)
        flash(f'{exam_type} mode activated', 'success')
    return redirect(url_for('index'))


# Regular Routes
@app.route('/regular/upload', methods=['GET', 'POST'])
def regular_upload():
    if request.method == 'POST':
        if session.get('exam_type') != 'Regular':
            return redirect(url_for('index'))

        semester = request.form.get('semester')
        month_year = request.form.get('month_year')
        file = request.files.get('master_file')
        elective_type = request.form.get('elective_type', 'general')

        if not all([semester, month_year, file]):
            flash('All fields required', 'error')
            return redirect(url_for('regular_upload'))

        filename = secure_filename(file.filename)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], 'temp', filename)
        os.makedirs(os.path.dirname(temp_path), exist_ok=True)
        file.save(temp_path)

        result = process_master_file(temp_path, 'Regular', semester, month_year, elective_type)
        os.remove(temp_path)

        if result['success']:
            flash('File uploaded successfully', 'success')
            session['last_upload'] = {'semester': semester, 'month_year': month_year, 'slots': result['slots']}
        else:
            flash(f'Error: {result["error"]}', 'error')

        return redirect(url_for('regular_upload'))

    return render_template('regular/upload.html', program_elective_available=PROGRAM_ELECTIVE_AVAILABLE,
                           open_elective_available=OPEN_ELECTIVE_AVAILABLE)


@app.route('/regular/allocation', methods=['GET', 'POST'])
def regular_allocation():
    if request.method == 'POST':
        if session.get('exam_type') != 'Regular':
            return redirect(url_for('index'))

        semester = request.form.get('semester')
        month_year = request.form.get('month_year')
        slot = request.form.get('slot')
        exam_date = request.form.get('exam_date')
        elective_type = request.form.get('elective_type', 'general')

        if not all([semester, month_year, slot]):
            flash('All fields required', 'error')
            return redirect(url_for('regular_allocation'))

        if not exam_date:
            flash('Please select an exam date', 'error')
            return redirect(url_for('regular_allocation'))

        possible_paths = []
        if elective_type != 'general':
            possible_paths.append(
                os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', elective_type, semester, month_year,
                             f'slot_{slot}'))
        possible_paths.append(
            os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', semester, month_year, f'slot_{slot}'))

        slot_folder = None
        for path in possible_paths:
            if os.path.exists(path):
                slot_folder = path
                break

        if not slot_folder:
            flash('Slot folder not found', 'error')
            return redirect(url_for('regular_allocation'))

        date_folder = os.path.join(slot_folder, exam_date)

        if not os.path.exists(date_folder):
            flash(f'Date folder {exam_date} not found', 'error')
            return redirect(url_for('regular_allocation'))

        slot_files = [f for f in os.listdir(date_folder) if f.endswith('.xlsx') and f.startswith(f'Slot_{slot}_')]

        if not slot_files:
            flash('Slot file not found', 'error')
            return redirect(url_for('regular_allocation'))

        slot_file_path = os.path.join(date_folder, slot_files[0])
        arrangement = generate_seating_arrangement(slot_file_path, date_folder, elective_type)

        if 'error' in arrangement:
            flash(f'Error: {arrangement["error"]}', 'error')
            return redirect(url_for('regular_allocation'))

        session['current_arrangement'] = arrangement
        session['arrangement_params'] = {'semester': semester, 'month_year': month_year, 'slot': slot,
                                         'exam_date': exam_date, 'elective_type': elective_type}

        return render_template('regular/results.html', arrangement=arrangement, params=session['arrangement_params'])

    return render_template('regular/allocation.html', program_elective_available=PROGRAM_ELECTIVE_AVAILABLE,
                           open_elective_available=OPEN_ELECTIVE_AVAILABLE)

@app.route('/regular/save_arrangement', methods=['POST'])
def save_arrangement():
    """
    Accept an edited rooms dict from the manual-edit UI and
    update the session arrangement + regenerate all output files.
    Payload: { "rooms": { "Room1": { "blocks": {...}, "total": N, "subjects": [...] }, ... } }
    """
    try:
        data = request.get_json()
        if not data or 'rooms' not in data:
            return jsonify({'success': False, 'error': 'No rooms data received'})

        # ── Retrieve the existing arrangement from session ─────────────────
        arrangement = session.get('current_arrangement')
        if not arrangement:
            return jsonify({'success': False, 'error': 'No arrangement in session'})

        params = session.get('arrangement_params', {})

        # ── Apply edits: replace blocks/total in the arrangement ──────────
        for room_name, room_payload in data['rooms'].items():
            if room_name not in arrangement['rooms']:
                continue
            arrangement['rooms'][room_name]['blocks'] = room_payload['blocks']
            arrangement['rooms'][room_name]['total']  = room_payload['total']
            if room_payload.get('subjects'):
                arrangement['rooms'][room_name]['subjects'] = room_payload['subjects']

        # Recompute top-level summary numbers
        room_totals = [arrangement['rooms'][r]['total'] for r in arrangement['rooms']]
        if room_totals:
            arrangement['summary']['average']            = sum(room_totals) / len(room_totals)
            arrangement['summary']['min']                = min(room_totals)
            arrangement['summary']['max']                = max(room_totals)
            arrangement['summary']['actual_difference']  = max(room_totals) - min(room_totals)
            arrangement['summary']['total_rooms']        = len(room_totals)

        # ── Locate the slot date folder from params ───────────────────────
        semester    = params.get('semester', '')
        month_year  = params.get('month_year', '')
        slot        = params.get('slot', '')
        exam_date   = params.get('exam_date', '')
        elective_type = params.get('elective_type', 'general')

        slot_date_folder = None
        possible_bases = []
        if elective_type != 'general':
            possible_bases.append(os.path.join(
                app.config['UPLOAD_FOLDER'], 'Regular', elective_type,
                semester, month_year, f'slot_{slot}', exam_date))
        possible_bases.append(os.path.join(
            app.config['UPLOAD_FOLDER'], 'Regular',
            semester, month_year, f'slot_{slot}', exam_date))

        for p in possible_bases:
            if os.path.exists(p):
                slot_date_folder = p
                break

        if not slot_date_folder:
            # Best effort: save JSON only
            print("WARNING: slot_date_folder not found; skipping file regeneration")
            session['current_arrangement'] = arrangement
            session.modified = True
            return jsonify({'success': True, 'warning': 'Files not regenerated (folder not found)'})

        # ── Find slot file path ───────────────────────────────────────────
        slot_files = [f for f in os.listdir(slot_date_folder)
                      if f.endswith('.xlsx') and f.startswith(f'Slot_{slot}_')]
        slot_file_path = os.path.join(slot_date_folder, slot_files[0]) if slot_files else None

        # ── Re-save arrangement JSON + TXT report ────────────────────────
        save_arrangement_files(arrangement, slot_date_folder,
                               slot_file_path or slot_date_folder)

        # ── Regenerate Excel ──────────────────────────────────────────────
        excel_output_path = os.path.join(slot_date_folder, 'seating_output.xlsx')
        if TXT_TO_EXCEL_AVAILABLE:
            try:
                txt_report_path = os.path.join(slot_date_folder, 'seating_report.txt')
                excel_result    = convert_txt_to_excel(txt_report_path, excel_output_path)
                if not (excel_result and excel_result.get('success')):
                    create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path)
            except Exception as e:
                print(f"Excel regen error: {e}")
                create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path)
        else:
            create_simple_excel_fallback(arrangement, slot_date_folder, excel_output_path)

        # ── Regenerate PDFs ───────────────────────────────────────────────
        if PDF_AVAILABLE:
            try:
                exam_date_display = exam_date.replace('-', '.')
                pdf_path = os.path.join(slot_date_folder, 'seating_arrangement.pdf')
                create_seating_pdf(
                    arrangement['rooms'],
                    arrangement['summary'],
                    pdf_path,
                    exam_date=exam_date_display,
                    semester=elective_type.capitalize() if elective_type != 'general' else 'Regular'
                )
                if 'qp_summary' in arrangement and arrangement['qp_summary']:
                    qp_pdf_path = os.path.join(slot_date_folder, 'qp_summary.pdf')
                    create_qp_summary_pdf(
                        arrangement['qp_summary'],
                        qp_pdf_path,
                        exam_date=exam_date_display
                    )
                # Regenerate master PDF
                try:
                    from txt_to_excel_converter import parse_txt, extract_metadata_from_txt
                    report_path = os.path.join(slot_date_folder, 'seating_report.txt')
                    with open(report_path, 'r', encoding='utf-8', errors='replace') as _f:
                        _txt = _f.read()
                    _metadata = extract_metadata_from_txt(_txt)
                    _meta, _rooms = parse_txt(_txt, _metadata)
                    _hall_map = get_room_to_hall_map()
                    master_pdf_path = os.path.join(slot_date_folder, 'master_list.pdf')
                    create_master_pdf(
                        _rooms, _meta, _hall_map, master_pdf_path,
                        institution=_metadata.get(
                            'institution',
                            'LBS INSTITUTE OF TECHNOLOGY FOR WOMEN, Poojappura, Thiruvananthapuram'),
                        exam_title=_metadata.get('exam_title', '')
                    )
                except Exception as _e:
                    print(f"Master PDF regen error: {_e}")
            except Exception as e:
                print(f"PDF regen error: {e}")

        # ── Update session ────────────────────────────────────────────────
        session['current_arrangement'] = arrangement
        session.modified = True

        print("✓ Arrangement saved and files regenerated after manual edit")
        return jsonify({'success': True})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@app.route('/regular/preview', methods=['GET', 'POST'])
def regular_preview():
    if request.method == 'POST':
        semester = request.form.get('semester')
        month_year = request.form.get('month_year')

        if not all([semester, month_year]):
            flash('All fields required', 'error')
            return redirect(url_for('regular_preview'))

        base_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', semester, month_year)
        master_file = os.path.join(base_path, 'master_list', 'Master_Sorted_List.xlsx')
        master_file_exists = os.path.exists(master_file)

        slot_folders = []
        excel_files_found = False
        pdf_files_found = False

        if os.path.exists(base_path):
            for item in os.listdir(base_path):
                if item.startswith('slot_') and os.path.isdir(os.path.join(base_path, item)):
                    slot = item.replace('slot_', '')
                    slot_folder = os.path.join(base_path, item)
                    date_folders = []

                    for date_item in os.listdir(slot_folder):
                        date_path = os.path.join(slot_folder, date_item)
                        if os.path.isdir(date_path):
                            files = []
                            for f in os.listdir(date_path):
                                file_type = 'other'
                                if f.endswith('.xlsx'):
                                    if f == 'seating_output.xlsx' or 'output' in f.lower():
                                        file_type = 'excel_output'
                                        excel_files_found = True
                                    else:
                                        file_type = 'xlsx'
                                elif f.endswith('.txt'):
                                    if 'qp_counts' in f:
                                        continue
                                    file_type = 'txt'
                                elif f.endswith('.json'):
                                    continue
                                elif f.endswith('.pdf'):
                                    if 'master_list' in f:
                                        file_type = 'pdf_master'
                                        pdf_files_found = True
                                    elif 'seating_arrangement' in f:
                                        file_type = 'pdf'
                                        pdf_files_found = True
                                    elif 'qp_summary' in f:
                                        file_type = 'pdf_qp'
                                        pdf_files_found = True
                                    else:
                                        file_type = 'pdf'
                                        pdf_files_found = True

                                if file_type != 'other' or (f.endswith('.xlsx') and 'Sorted_List' in f):
                                    file_path = os.path.join(date_path, f)
                                    file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                                    file_modified = datetime.fromtimestamp(
                                        os.path.getmtime(file_path)) if os.path.exists(
                                        file_path) else datetime.now()

                                    files.append({
                                        'name': f,
                                        'type': file_type,
                                        'size': file_size,
                                        'modified': file_modified.strftime('%Y-%m-%d %H:%M:%S')
                                    })

                            files.sort(key=lambda x: (
                                0 if x['type'] == 'pdf_master' else
                                1 if x['type'] == 'pdf' else
                                2 if x['type'] == 'pdf_qp' else
                                3 if x['type'] == 'excel_output' else
                                4 if x['type'] == 'xlsx' and 'Sorted_List' in x['name'] else
                                5, x['name']))

                            if files:
                                date_folders.append({'date': date_item, 'files': files})

                    if date_folders:
                        slot_folders.append({'name': slot, 'date_folders': date_folders})

        return render_template('regular/preview.html', semester=semester, month_year=month_year,
                               master_file_exists=master_file_exists, slot_folders=slot_folders,
                               excel_files_found=excel_files_found, pdf_files_found=pdf_files_found)

    return render_template('regular/preview.html')


@app.route('/view_excel/<exam_type>/<semester>/<month_year>/<slot>/<date>/<filename>')
def view_excel(exam_type, semester, month_year, slot, date, filename):
    """Route to view Excel file in browser"""
    paths = [
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, semester, month_year, f'slot_{slot}', date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'program_elective', semester, month_year, f'slot_{slot}',
                     date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'open_elective', semester, month_year, f'slot_{slot}',
                     date, filename)
    ]

    for path in paths:
        if os.path.exists(path):
            try:
                import pandas as pd
                excel_file = pd.ExcelFile(path)
                sheet_names = excel_file.sheet_names

                df_dict = {}
                for sheet in sheet_names:
                    df_dict[sheet] = pd.read_excel(path, sheet_name=sheet)

                metadata = {
                    'exam_date': date.replace('-', '.'),
                    'filename': filename,
                    'semester': semester,
                    'slot': slot,
                    'month_year': month_year
                }

                return render_template('excel_viewer.html',
                                       df_dict=df_dict,
                                       filename=filename,
                                       metadata=metadata,
                                       sheet_names=sheet_names)
            except Exception as e:
                print(f"Error viewing Excel: {e}")
                return send_file(path, as_attachment=False)

    flash('File not found', 'error')
    return redirect(url_for('regular_preview'))


@app.route('/view_pdf/<exam_type>/<semester>/<month_year>/<slot>/<date>/<filename>')
def view_pdf(exam_type, semester, month_year, slot, date, filename):
    """Route to view PDF file in browser"""
    paths = [
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, semester, month_year, f'slot_{slot}', date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'program_elective', semester, month_year, f'slot_{slot}',
                     date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'open_elective', semester, month_year, f'slot_{slot}',
                     date, filename)
    ]

    for path in paths:
        if os.path.exists(path):
            return send_file(path, as_attachment=False, mimetype='application/pdf')

    flash('PDF file not found', 'error')
    return redirect(url_for('regular_preview'))


# ── Series Routes ─────────────────────────────────────────────────────────────

@app.route('/series/entry', methods=['GET', 'POST'])
def series_entry():
    if request.method == 'POST':
        if session.get('exam_type') != 'Series':
            return jsonify({'success': False, 'error': 'Invalid exam type'})

        data = request.get_json()
        rooms = data.get('rooms', {})
        classes = data.get('classes', [])
        subjects = data.get('subjects', {})

        total_students = sum(cls.get('strength', 0) for cls in classes)

        # ── FIXED: rooms values are now per-block dicts, not plain ints ──────
        total_capacity = sum(
            sum(v.values()) if isinstance(v, dict) else int(v)
            for v in rooms.values()
        )
        # ─────────────────────────────────────────────────────────────────────

        session['series_data'] = {
            'rooms': rooms,
            'classes': classes,
            'subjects': subjects,
            'total_students': total_students,
            'total_capacity': total_capacity
        }

        return jsonify({'success': True})

    return render_template('series/entry.html')


@app.route('/series/allocation', methods=['GET', 'POST'])
def series_allocation():
    if request.method == 'POST':
        if session.get('exam_type') != 'Series' or 'series_data' not in session:
            flash('Please select Series exam and enter data first', 'error')
            return redirect(url_for('index'))

        algorithm = request.form.get('algorithm', 'general')

        try:
            arrangement = generate_series_arrangement(session['series_data'])

            if 'error' in arrangement:
                flash(f'Error: {arrangement["error"]}', 'error')
                return redirect(url_for('series_allocation'))

            session['series_arrangement'] = arrangement
            session.modified = True

            flash('Seating arrangement generated successfully!', 'success')
            return redirect(url_for('series_results'))

        except Exception as e:
            print(f"\n!!! EXCEPTION IN SERIES ALLOCATION !!!")
            import traceback
            traceback.print_exc()
            flash(f'Error generating arrangement: {str(e)}', 'error')
            return redirect(url_for('series_allocation'))

    return render_template('series/allocation.html')


@app.route('/series/results')
def series_results():
    if 'series_arrangement' not in session:
        flash('No arrangement found. Please generate one first.', 'warning')
        return redirect(url_for('series_allocation'))

    return render_template('series/results.html', arrangement=session['series_arrangement'])


@app.route('/series/files')
def series_files():
    """View all generated series exam output files, grouped by date."""
    series_base = os.path.join(app.config['UPLOAD_FOLDER'], 'Series')
    sessions = []

    if os.path.exists(series_base):
        for date_folder in sorted(os.listdir(series_base), reverse=True):
            date_path = os.path.join(series_base, date_folder)
            if not os.path.isdir(date_path):
                continue

            files = []
            for fname in sorted(os.listdir(date_path)):
                fpath = os.path.join(date_path, fname)
                if not os.path.isfile(fpath):
                    continue
                ext = fname.rsplit('.', 1)[-1].lower() if '.' in fname else ''
                ftype = 'excel' if ext == 'xlsx' else ('pdf' if ext == 'pdf' else 'other')
                files.append({
                    'name': fname,
                    'type': ftype,
                    'size': os.path.getsize(fpath),
                    'modified': datetime.fromtimestamp(os.path.getmtime(fpath)).strftime('%Y-%m-%d %H:%M'),
                    'date_folder': date_folder,
                })

            if files:
                sessions.append({'date': date_folder, 'files': files})

    return render_template('series/files.html', sessions=sessions)


@app.route('/series/download/<date_folder>/<filename>')
def series_download_file(date_folder, filename):
    """Download or inline-view a series output file."""
    safe_date = os.path.basename(date_folder)
    safe_file = os.path.basename(filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Series', safe_date, safe_file)

    if not os.path.exists(file_path):
        flash('File not found.', 'error')
        return redirect(url_for('series_files'))

    ext = safe_file.rsplit('.', 1)[-1].lower() if '.' in safe_file else ''
    if ext == 'pdf':
        return send_file(file_path, as_attachment=False, mimetype='application/pdf')
    else:
        return send_file(file_path, as_attachment=True)


# ── Download Routes ───────────────────────────────────────────────────────────

@app.route('/download_master/<exam_type>/<semester>/<month_year>')
def download_master(exam_type, semester, month_year):
    paths = [
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, semester, month_year, 'master_list',
                     'Master_Sorted_List.xlsx'),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'program_elective', semester, month_year, 'master_list',
                     'Master_Sorted_List.xlsx'),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'open_elective', semester, month_year, 'master_list',
                     'Master_Sorted_List.xlsx')
    ]
    for path in paths:
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    flash('File not found', 'error')
    return redirect(url_for('regular_preview'))


@app.route('/download_file/<exam_type>/<semester>/<month_year>/<slot>/<date>/<filename>')
def download_file(exam_type, semester, month_year, slot, date, filename):
    paths = [
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, semester, month_year, f'slot_{slot}', date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'program_elective', semester, month_year, f'slot_{slot}',
                     date, filename),
        os.path.join(app.config['UPLOAD_FOLDER'], exam_type, 'open_elective', semester, month_year, f'slot_{slot}',
                     date, filename)
    ]
    for path in paths:
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    flash('File not found', 'error')
    return redirect(url_for('regular_preview'))


@app.route('/api/available-dates')
def api_available_dates():
    """API endpoint to fetch available dates for a given semester, month-year, and slot"""
    semester = request.args.get('semester')
    month_year = request.args.get('month_year')
    slot = request.args.get('slot')

    if not all([semester, month_year, slot]):
        return jsonify({'dates': [], 'error': 'Missing parameters'})

    possible_paths = [
        os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', semester, month_year, f'slot_{slot}'),
        os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', 'program_elective', semester, month_year, f'slot_{slot}'),
        os.path.join(app.config['UPLOAD_FOLDER'], 'Regular', 'open_elective', semester, month_year, f'slot_{slot}')
    ]

    dates = []
    for path in possible_paths:
        if os.path.exists(path) and os.path.isdir(path):
            for item in os.listdir(path):
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    slot_files = [f for f in os.listdir(item_path)
                                  if f.endswith('.xlsx') and f.startswith(f'Slot_{slot}_')]
                    if slot_files:
                        dates.append(item)

    dates = sorted(list(set(dates)))
    return jsonify({'dates': dates})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
