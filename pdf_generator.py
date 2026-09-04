import os
import re
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping

# Hall number mapping for rooms
HALL_NUMBERS = {
    "Room1": "108", "Room2": "109", "Room3": "111", "Room4": "112",
    "Room5": "208", "Room6": "209", "Room7": "211", "Room8": "212",
    "Room9": "308", "Room10": "309", "Room11": "311", "Room12": "312",
    # Additional rooms with sequential numbering
    "Room13": "313", "Room14": "314", "Room15": "315", "Room16": "316",
    "Room17": "317", "Room18": "318", "Room19": "319", "Room20": "320",
    "Room21": "321", "Room22": "322", "Room23": "323", "Room24": "324",
    "Room25": "325", "Room26": "326", "Room27": "327", "Room28": "328",
    "Room29": "329", "Room30": "330"
}


def get_hall_number(room_name):
    """Get hall number for a given room name"""
    match = re.search(r'Room(\d+)', room_name)
    if match:
        room_num = int(match.group(1))
        if room_name in HALL_NUMBERS:
            return HALL_NUMBERS[room_name]
        else:
            base_hall = 330
            return str(base_hall + (room_num - 30))
    return room_name


def create_seating_pdf(rooms_data, summary_data, output_path, exam_date=None, semester=None):
    """
    Create a professional PDF seating arrangement report

    Args:
        rooms_data: Dictionary of room data from arrangement['rooms']
        summary_data: Dictionary of summary data from arrangement['summary']
        output_path: Path where PDF will be saved
        exam_date: Exam date for header
        semester: Semester for header
    """
    try:
        doc = SimpleDocTemplate(
            output_path,
            pagesize=landscape(A4),
            rightMargin=0.5 * cm,
            leftMargin=0.5 * cm,
            topMargin=1 * cm,
            bottomMargin=1 * cm
        )

        styles = getSampleStyleSheet()
        title_style = styles['Heading1']
        title_style.alignment = 1
        title_style.fontSize = 16
        title_style.spaceAfter = 12

        heading_style = styles['Heading2']
        heading_style.fontSize = 14
        heading_style.spaceAfter = 8
        heading_style.spaceBefore = 12

        normal_style = styles['Normal']
        normal_style.fontSize = 10

        elements = []

        header_text = f"LBS INSTITUTE OF TECHNOLOGY FOR WOMEN, POOJAPPURA"
        elements.append(Paragraph(header_text, title_style))

        if exam_date:
            date_text = f"EXAM DATE: {exam_date}"
            elements.append(Paragraph(date_text, heading_style))

        if semester:
            sem_text = f"SEMESTER: {semester}"
            elements.append(Paragraph(sem_text, normal_style))

        elements.append(Spacer(1, 0.5 * cm))

        elements.append(Paragraph("SUMMARY", heading_style))

        summary_table_data = [
            ["Total Students", "Total Rooms", "Average/Room", "Min", "Max", "Difference"],
            [
                str(summary_data.get('total_rooms_placed', summary_data.get('student_count', 0))),
                str(summary_data.get('total_rooms', 0)),
                f"{summary_data.get('average', 0):.1f}",
                str(summary_data.get('min', 0)),
                str(summary_data.get('max', 0)),
                str(summary_data.get('actual_difference', 0))
            ]
        ]

        summary_table = Table(summary_table_data, colWidths=[2.5 * cm] * 6)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))

        elements.append(summary_table)
        elements.append(Spacer(1, 1 * cm))

        elements.append(Paragraph("SEATING ARRANGEMENT - ROOM WISE", heading_style))
        elements.append(Spacer(1, 0.5 * cm))

        sorted_rooms = sorted(rooms_data.keys(), key=lambda x: int(re.search(r'\d+', x).group()))

        for room_name in sorted_rooms:
            room_data = rooms_data[room_name]
            hall_number = get_hall_number(room_name)

            room_header = f"{room_name} (Hall No: {hall_number}) - Total: {room_data['total']} Students"
            elements.append(Paragraph(room_header, styles['Heading3']))

            block_order = ['Left1', 'Left3', 'Middle2', 'Right1', 'Right3']
            block_to_column = {
                'Left1': '1', 'Left3': '2', 'Middle2': '3',
                'Right1': '4', 'Right3': '5'
            }

            blocks = []
            max_rows = 0

            for block_name in block_order:
                if block_name in room_data['blocks']:
                    block = room_data['blocks'][block_name]
                    students = block.get('students', [])
                    blocks.append({
                        'name': block_name,
                        'column_num': block_to_column[block_name],
                        'students': students,
                        'capacity': block.get('capacity', 7),
                        'count': block.get('count', 0)
                    })
                    max_rows = max(max_rows, len(students))

            header_row = ['Row']
            for block in blocks:
                header_row.append(f"Col {block['column_num']}\n({block['count']}/{block['capacity']})")

            table_data = [header_row]

            for row_num in range(max_rows):
                row_data = [str(row_num + 1)]
                for block in blocks:
                    if row_num < len(block['students']):
                        student = block['students'][row_num]
                        if '(' in student:
                            match = re.search(r'\(([^)]+)\)', student)
                            if match:
                                row_data.append(match.group(1))
                            else:
                                row_data.append(student[:12])
                        else:
                            row_data.append(student[:12])
                    else:
                        row_data.append('--')
                table_data.append(row_data)

            col_widths = [1.2 * cm]
            for _ in blocks:
                col_widths.append(2.8 * cm)

            table = Table(table_data, colWidths=col_widths)

            table_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]

            for i in range(1, len(table_data)):
                if i % 2 == 0:
                    table_style.append(('BACKGROUND', (0, i), (-1, i), colors.lightgrey))

            table.setStyle(TableStyle(table_style))

            elements.append(table)
            elements.append(Spacer(1, 0.3 * cm))

            if room_data.get('subjects'):
                subjects_text = "Subjects: " + ", ".join(room_data['subjects'])
                elements.append(Paragraph(subjects_text, normal_style))

            elements.append(Spacer(1, 0.8 * cm))

            if sorted_rooms.index(room_name) % 2 == 1 and sorted_rooms.index(room_name) < len(sorted_rooms) - 1:
                elements.append(PageBreak())

        doc.build(elements)
        return True

    except Exception as e:
        print(f"Error creating PDF: {e}")
        import traceback
        traceback.print_exc()
        return False


def create_qp_summary_pdf(qp_data, output_path, exam_date=None):
    """
    Create a professional PDF for QP (Question Paper) summary

    Args:
        qp_data: Dictionary containing room_wise and subject_summary
        output_path: Path where PDF will be saved
        exam_date: Exam date for header
    """
    try:
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=1 * cm,
            leftMargin=1 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm
        )

        styles = getSampleStyleSheet()
        title_style = styles['Heading1']
        title_style.alignment = 1
        title_style.fontSize = 16
        title_style.spaceAfter = 12

        heading_style = styles['Heading2']
        heading_style.fontSize = 14
        heading_style.spaceAfter = 8

        subheading_style = styles['Heading3']
        subheading_style.fontSize = 12
        subheading_style.spaceAfter = 6

        normal_style = styles['Normal']
        normal_style.fontSize = 10

        elements = []

        elements.append(Paragraph("QUESTION PAPER COUNT SUMMARY", title_style))

        if exam_date:
            elements.append(Paragraph(f"Exam Date: {exam_date}", heading_style))

        elements.append(Spacer(1, 0.5 * cm))

        elements.append(Paragraph("ROOM-WISE SUBJECT DISTRIBUTION", subheading_style))
        elements.append(Spacer(1, 0.3 * cm))

        if qp_data.get('room_wise'):
            room_groups = {}
            for item in qp_data['room_wise']:
                room = item['Room']
                if room not in room_groups:
                    room_groups[room] = []
                room_groups[room].append(item)

            sorted_rooms = sorted(room_groups.keys(), key=lambda x: int(re.search(r'\d+', x).group()))

            for room in sorted_rooms:
                hall_number = get_hall_number(room)
                elements.append(Paragraph(f"{room} (Hall No: {hall_number})", subheading_style))

                room_items = room_groups[room]
                table_data = [['Subject', 'Count']]

                for item in room_items:
                    table_data.append([item['Subject'], str(item['Student Count'])])

                room_total = sum(item['Student Count'] for item in room_items)
                table_data.append(['TOTAL', str(room_total)])

                table = Table(table_data, colWidths=[12 * cm, 3 * cm])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -2), colors.white),
                    ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                    ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))

                elements.append(table)
                elements.append(Spacer(1, 0.5 * cm))

        elements.append(PageBreak())
        elements.append(Paragraph("SUBJECT SUMMARY", subheading_style))
        elements.append(Spacer(1, 0.3 * cm))

        if qp_data.get('subject_summary'):
            subject_data = [['Subject', 'Total Students']]
            for subject, count in sorted(qp_data['subject_summary'].items()):
                subject_data.append([subject, str(count)])

            subject_data.append(['GRAND TOTAL', str(qp_data.get('total_students', 0))])

            table = Table(subject_data, colWidths=[12 * cm, 3 * cm])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -2), colors.white),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))

            elements.append(table)

        doc.build(elements)
        return True

    except Exception as e:
        print(f"Error creating QP summary PDF: {e}")
        return False


def create_master_pdf(rooms, meta, hall_map, output_path,
                      institution=None, exam_title=None):
    """
    Generate a PDF version of the MASTER sheet.

    Columns:  BRANCH  |  REGISTER NUMBERS  |  HALL NO

    Args:
        rooms       - list of room dicts (each has 'room_name' and 'rows')
        meta        - {'subjects': {...}} dict (may be empty)
        hall_map    - {'Room1': '108', ...}
        output_path - path to write the PDF
        institution - institution name for header (uses default if None)
        exam_title  - exam title for sub-header (omitted if None)
    Returns True on success, False on failure.
    """
    try:
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.styles import ParagraphStyle
        from collections import defaultdict

        # ── branch code -> department mapping ────────────────────────────────
        _BM = {
            "CS": "CSE", "EC": "ECE", "CE": "CE", "IT": "IT",
            "ME": "ME",  "EE": "EEE", "AE": "AE", "ER": "ERE", "ECE": "ERE"
        }

        def _branch(reg):
            if not reg: return None
            m = re.search(r"[A-Z]{3}\d{2}([A-Z]+)\d+$", reg.upper())
            return _BM.get(m.group(1)) if m else None

        def _supply(reg):
            return bool(reg and reg.upper().startswith("LL"))

        def _range_str(regs):
            if not regs: return ""
            cl = sorted(set(regs))
            if len(cl) == 1: return cl[0] + " (1)"
            groups, g = [], [cl[0]]
            for reg in cl[1:]:
                pm  = re.search(r"\d+$", g[-1])
                cm_ = re.search(r"\d+$", reg)
                if (pm and cm_ and g[-1][:pm.start()] == reg[:cm_.start()]
                        and int(cm_.group()) == int(pm.group()) + 1):
                    g.append(reg)
                else:
                    groups.append(g)
                    g = [reg]
            groups.append(g)
            parts = []
            for grp in groups:
                nm = re.search(r"\d+$", grp[-1])
                parts.append(grp[0] + "-" + nm.group()
                             if len(grp) > 1 and nm else grp[0])
            return ", ".join(parts) + f" ({len(cl)})"

        DEPT_ORDER = ["CE", "CSE", "ECE", "IT", "ERE", "ME", "EEE", "AE"]

        # ── aggregate students: branch -> hall -> {regular, supply} ──────────
        bh = defaultdict(lambda: defaultdict(lambda: {"r": [], "s": []}))
        for room in rooms:
            hall = hall_map.get(room["room_name"], room["room_name"])
            for _, vals in room.get("rows", []):
                for v in vals:
                    if v:
                        br = _branch(v)
                        if br:
                            key = "s" if _supply(v) else "r"
                            bh[br][hall][key].append(v)
        for br in bh:
            for h in bh[br]:
                bh[br][h]["r"] = sorted(set(bh[br][h]["r"]))
                bh[br][h]["s"] = sorted(set(bh[br][h]["s"]))

        ordered_depts = ([b for b in DEPT_ORDER if b in bh] +
                         [b for b in sorted(bh) if b not in DEPT_ORDER])

        # ── PDF document ──────────────────────────────────────────────────────
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=1.5 * cm, leftMargin=1.5 * cm,
            topMargin=1.5 * cm,  bottomMargin=1.5 * cm,
        )

        def ps(name, size, bold=False, align=TA_CENTER, color=colors.black,
               space_before=0, space_after=4):
            return ParagraphStyle(
                name,
                fontName="Helvetica-Bold" if bold else "Helvetica",
                fontSize=size, textColor=color, alignment=align,
                spaceBefore=space_before, spaceAfter=space_after,
            )

        hdr1_style = ps("MH1", 13, bold=True, space_after=3)
        hdr2_style = ps("MH2", 11, bold=True, space_after=3)
        hdr3_style = ps("MH3", 10, bold=True, space_after=10)
        cell_c     = ps("MCC",  9, bold=True,  align=TA_CENTER)
        cell_l     = ps("MCL",  9, bold=True,  align=TA_LEFT)
        col_hdr    = ps("MCH", 10, bold=True,  align=TA_CENTER,
                        color=colors.white)

        elems = []

        inst = (institution or
                "LBS INSTITUTE OF TECHNOLOGY FOR WOMEN, "
                "Poojappura, Thiruvananthapuram")
        elems.append(Paragraph(inst, hdr1_style))
        elems.append(Paragraph(
            "APJ ABDUL KALAM TECHNOLOGICAL UNIVERSITY (APJAKTU)", hdr2_style))
        if exam_title:
            elems.append(Paragraph(exam_title, hdr3_style))
        else:
            elems.append(Spacer(1, 0.3 * cm))

        # ── table ─────────────────────────────────────────────────────────────
        # Page usable width ≈ A4(21cm) − margins(3cm) = 18 cm
        COL_W = [3 * cm, 12 * cm, 3 * cm]

        table_data = [[
            Paragraph("BRANCH",           col_hdr),
            Paragraph("REGISTER NUMBERS", col_hdr),
            Paragraph("HALL NO",          col_hdr),
        ]]

        ALT_A = colors.HexColor("#EBF5FB")   # light blue for odd dept groups
        ALT_B = colors.white                  # white for even dept groups
        row_fills = []
        dept_group_idx = 0

        for br in ordered_depts:
            halls = sorted(bh[br].keys(),
                           key=lambda x: int(x) if x.isdigit() else 999)
            br_rows = [
                (h, _range_str(bh[br][h]["r"] + bh[br][h]["s"]))
                for h in halls
                if bh[br][h]["r"] or bh[br][h]["s"]
            ]
            if not br_rows:
                continue

            fill = ALT_A if dept_group_idx % 2 == 0 else ALT_B
            dept_group_idx += 1

            for i, (hall, rng) in enumerate(br_rows):
                br_para   = Paragraph(br,  cell_c) if i == 0 else Paragraph("", cell_c)
                rng_para  = Paragraph(rng, cell_l)
                hall_para = Paragraph(str(hall), cell_c)
                table_data.append([br_para, rng_para, hall_para])
                row_fills.append((len(table_data) - 1, fill))

        style_cmds = [
            ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#1A5276")),
            ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
            ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",      (0, 0), (-1, 0),  10),
            ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#AEB6BF")),
            ("BOX",           (0, 0), (-1, -1), 1.2, colors.HexColor("#1A5276")),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("LINEBELOW",     (0, 0), (-1, 0),  1.5, colors.HexColor("#1A5276")),
        ]
        for ridx, fill in row_fills:
            style_cmds.append(("BACKGROUND", (0, ridx), (-1, ridx), fill))

        tbl = Table(table_data, colWidths=COL_W, repeatRows=1)
        tbl.setStyle(TableStyle(style_cmds))
        elems.append(tbl)

        doc.build(elems)
        return True

    except Exception as e:
        import traceback
        print(f"Error creating master PDF: {e}")
        traceback.print_exc()
        return False
