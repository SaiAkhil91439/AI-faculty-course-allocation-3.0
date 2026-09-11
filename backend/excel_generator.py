import os
from typing import Dict, Any, List
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_allocation_excel(
    allocation_res: Dict[str, Any],
    output_path: str = "output/faculty_allocation.xlsx"
) -> str:
    """
    Generates a professionally styled Excel workbook matching the exact target workload layout.
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = openpyxl.Workbook()

    # -------------------------------------------------------------
    # TAB 1: Faculty Workload & Allocation
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Faculty Workload Allocation"

    # Title Banner
    ws1.merge_cells("A1:V1")
    title_cell = ws1["A1"]
    title_cell.value = "FACULTY WORKLOAD & COURSE ALLOCATION REPORT"
    title_cell.font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 40

    # Subtitle Info
    ws1.merge_cells("A2:V2")
    sub_cell = ws1["A2"]
    sub_cell.value = f"Status: {allocation_res['status']} | Average Compatibility Score: {allocation_res['average_score']}/100 | Total Sections: {allocation_res['num_sections']} | Subjects: {allocation_res['num_subjects']}"
    sub_cell.font = Font(name="Arial", size=11, italic=True, color="1F4E79")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 25

    # Header Row 3 & 4
    headers = [
        "Sl. No.", "Name of the Faculty", "Designation", "Additional duties", "Present Workload (hrs/wk)",
        "Program", "Year", "Section", "Name of the COURSE", "Course Code", "L", "T", "P", "C",
        "No. of section", "No. of students", "Total L HOURS", "Total T HOURS", "Total P HOURS",
        "Total hours course/wk", "Total workload/wk", "Comments / Role"
    ]

    ws1.row_dimensions[4].height = 30
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    for col_idx, text in enumerate(headers, 1):
        cell = ws1.cell(row=4, column=col_idx, value=text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Data Rows
    allocs = allocation_res["faculty_allocations"]
    row_cursor = 5
    sl_no = 1

    # Group allocations by Lecture Faculty
    fac_grouped = {}
    for a in allocs:
        f_name = a["Lecture_Faculty_Name"]
        if f_name not in fac_grouped:
            fac_grouped[f_name] = []
        fac_grouped[f_name].append(a)

    alt_fill_1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    alt_fill_2 = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    dean_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft green for Dean/HOD

    for f_name, f_allocs in fac_grouped.items():
        first_alloc = f_allocs[0]
        desig = first_alloc.get("Lecture_Designation", "Professor")
        role = first_alloc.get("Lecture_Role", "Faculty")

        # Determine workload hours
        w_hours = "8-10" if role == "Dean" else ("12-14" if role == "HOD" else "14-16")
        row_fill = dean_fill if role in ("Dean", "HOD") else (alt_fill_1 if sl_no % 2 != 0 else alt_fill_2)

        start_row = row_cursor
        for a in f_allocs:
            secs = a["Sections"].split(",")
            is_high_exp = a.get("Is_High_Experience_Split", False)
            exp_years = a.get("Experience_Years", 0.0)

            comment_str = f"High Exp Split ({a['Num_Sections']} Secs, {exp_years} yrs exp)" if is_high_exp else f"Main Lecturer ({role})"

            for sec in secs:
                sec_clean = sec.strip()
                ws1.cell(row=row_cursor, column=1, value=sl_no).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=2, value=f_name)
                ws1.cell(row=row_cursor, column=3, value=desig)
                ws1.cell(row=row_cursor, column=4, value=role)
                ws1.cell(row=row_cursor, column=5, value=w_hours).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=6, value="AIML").alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=7, value="III").alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=8, value=sec_clean).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=9, value=a["Subject"])
                ws1.cell(row=row_cursor, column=10, value=f"CS{100+sl_no}").alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=11, value=a.get("L", 3)).alignment = Alignment(horizontal="center") # L
                ws1.cell(row=row_cursor, column=12, value=a.get("T", 0)).alignment = Alignment(horizontal="center") # T
                ws1.cell(row=row_cursor, column=13, value=a.get("P", 2)).alignment = Alignment(horizontal="center") # P
                ws1.cell(row=row_cursor, column=14, value=a.get("C", 4)).alignment = Alignment(horizontal="center") # C
                ws1.cell(row=row_cursor, column=15, value=1).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=16, value=65).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=17, value=a.get("L", 3)).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=18, value=a.get("T", 0)).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=19, value=a.get("P", 2)).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=20, value=a.get("Total_Course_Hours", 5)).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=21, value=a.get("Faculty_Total_Weekly_Hours", 15)).alignment = Alignment(horizontal="center")
                ws1.cell(row=row_cursor, column=22, value=comment_str)

                for col in range(1, 23):
                    cell = ws1.cell(row=row_cursor, column=col)
                    cell.fill = row_fill
                    cell.border = thin_border
                    cell.font = Font(name="Arial", size=9)

                row_cursor += 1

        sl_no += 1

    # Adjust Column Widths
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # -------------------------------------------------------------
    # TAB 2: Course Component Schedule (L & T/P Breakdown)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Course Section Schedule")

    ws2.merge_cells("A1:H1")
    t2_cell = ws2["A1"]
    t2_cell.value = "COURSE SECTION FACULTY ASSIGNMENTS (L = 1 Faculty, T/P = 2 Faculty)"
    t2_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    t2_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    t2_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 35

    s2_headers = [
        "Course / Subject", "L-T-P-C", "Section", "Lecture (L) Faculty (1)",
        "Teaching Experience (Yrs)", "Tutorial / Practical (T/P) Faculty 1", "Tutorial / Practical (T/P) Faculty 2", "Score"
    ]

    for col_idx, text in enumerate(s2_headers, 1):
        cell = ws2.cell(row=3, column=col_idx, value=text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    r2_idx = 4
    for a in allocs:
        is_high = a.get("Is_High_Experience_Split", False)
        fac_label = f"{a['Lecture_Faculty_Name']} ({a['Lecture_Designation']})" + (" [High Exp 2+1 Split]" if is_high else "")
        ws2.cell(row=r2_idx, column=1, value=a["Subject"])
        ws2.cell(row=r2_idx, column=2, value=a.get("L-T-P-C", "3-0-2-4")).alignment = Alignment(horizontal="center")
        ws2.cell(row=r2_idx, column=3, value=a["Sections"]).alignment = Alignment(horizontal="center")
        ws2.cell(row=r2_idx, column=4, value=fac_label)
        ws2.cell(row=r2_idx, column=5, value=a.get("Experience_Years", 0.0)).alignment = Alignment(horizontal="center")
        ws2.cell(row=r2_idx, column=6, value=a["TP_Faculty_1_Name"])
        ws2.cell(row=r2_idx, column=7, value=a["TP_Faculty_2_Name"])
        ws2.cell(row=r2_idx, column=8, value=a["Score"]).alignment = Alignment(horizontal="center")

        for col in range(1, 9):
            cell = ws2.cell(row=r2_idx, column=col)
            cell.border = thin_border
            cell.font = Font(name="Arial", size=9.5)
        r2_idx += 1


    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(output_path)
    return output_path

    

