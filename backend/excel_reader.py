import pandas as pd
from typing import Union, BinaryIO, Tuple, Optional

def derive_skill_level(exp_years: float) -> str:
    if exp_years >= 8:
        return "Expert"
    elif exp_years >= 5:
        return "Advanced"
    elif exp_years >= 2:
        return "Intermediate"
    else:
        return "Beginner"

def extract_role(designation: str, addl_duties: str, fac_name: str) -> str:
    """Detects role: Dean, HOD, Associate Professor, Assistant Professor, Professor."""
    combined = f"{designation} {addl_duties} {fac_name}".lower()
    if "dean" in combined:
        return "Dean"
    elif "hod" in combined or "head of" in combined or "h.o.d" in combined:
        return "HOD"
    elif "associate" in combined or "assoc" in combined:
        return "Associate Professor"
    elif "assistant" in combined or "asst" in combined:
        return "Assistant Professor"
    elif "professor" in combined or "prof" in combined:
        return "Professor"
    return "Faculty"

def parse_ltpc_string(ltpc_str: str) -> Tuple[int, int, int, int]:
    """Parses 'L-T-P-C' format string like '3-0-2-4' or '1-2-0-2' into (L, T, P, C)."""
    try:
        parts = [int(x.strip()) for x in str(ltpc_str).split("-")]
        if len(parts) >= 4:
            return parts[0], parts[1], parts[2], parts[3]
        elif len(parts) == 3:
            return parts[0], parts[1], parts[2], parts[0] + parts[1] + parts[2]
    except Exception:
        pass
    return 3, 0, 2, 4

def read_course_details(file_source: Union[str, BinaryIO]) -> pd.DataFrame:
    """
    Reads Course Details Excel file containing Courses and L-T-P-C columns.
    Returns cleaned DataFrame with columns: ['Subject', 'L', 'T', 'P', 'C', 'L-T-P-C', 'Total_Hours'].
    """
    df = pd.read_excel(file_source)
    df.columns = [str(col).strip() for col in df.columns]

    course_col = None
    for c in ["Courses", "Course", "Subject", "Course Name", "Name of the COURSE", "Course_Name"]:
        if c in df.columns:
            course_col = c
            break
    if not course_col:
        course_col = df.columns[0]

    ltpc_col = None
    for c in ["L-T-P-C", "L-T-P-C Structure", "LTPC", "Structure"]:
        if c in df.columns:
            ltpc_col = c
            break

    records = []
    for _, row in df.iterrows():
        subj = str(row[course_col]).strip()
        if not subj or subj.lower() in ('nan', 'none', ''):
            continue

        l_val, t_val, p_val, c_val = 3, 0, 2, 4
        if ltpc_col and pd.notna(row[ltpc_col]):
            l_val, t_val, p_val, c_val = parse_ltpc_string(row[ltpc_col])
        else:
            l_val = int(pd.to_numeric(row.get("L", 3), errors='coerce') or 3)
            t_val = int(pd.to_numeric(row.get("T", 0), errors='coerce') or 0)
            p_val = int(pd.to_numeric(row.get("P", 2), errors='coerce') or 2)
            c_val = int(pd.to_numeric(row.get("C", 4), errors='coerce') or 4)

        total_hrs = l_val + t_val + p_val
        records.append({
            "Subject": subj,
            "L": l_val,
            "T": t_val,
            "P": p_val,
            "C": c_val,
            "L-T-P-C": f"{l_val}-{t_val}-{p_val}-{c_val}",
            "Total_Hours": total_hrs
        })

    return pd.DataFrame(records)


def parse_wide_combined_sheet(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parses a single wide-format Excel sheet into standard Preferences and Experience DataFrames.
    """
    df.columns = [str(col).strip() for col in df.columns]
    
    # Identify faculty name column
    name_col = None
    for c in ["Faculty Name", "Faculty_Name", "FacultyName", "Name"]:
        if c in df.columns:
            name_col = c
            break
    if not name_col:
        name_col = df.columns[0]

    desig_col = None
    for c in ["Designation", "Role", "Position"]:
        if c in df.columns:
            desig_col = c
            break

    addl_col = None
    for c in ["Additional Duties", "Additional_Duties", "Addl Duties", "Addl_Duties"]:
        if c in df.columns:
            addl_col = c
            break

    qual_col = None
    for c in ["Qualification", "Degree", "Total Experience"]:
        if c in df.columns:
            qual_col = c
            break

    pref_rows = []
    exp_rows = []

    for idx, row in df.iterrows():
        fac_id = f"F{idx+1:03d}"
        fac_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else f"Faculty_{idx+1}"
        desig = str(row[desig_col]).strip() if desig_col and pd.notna(row[desig_col]) else "Professor"
        addl = str(row[addl_col]).strip() if addl_col and pd.notna(row[addl_col]) else "None"
        qual = str(row[qual_col]).strip() if qual_col and pd.notna(row[qual_col]) else "PhD CSE"

        role = extract_role(desig, addl, fac_name)

        p_dict = {
            "Faculty_ID": fac_id,
            "Faculty_Name": fac_name,
            "Designation": desig,
            "Additional_Duties": addl,
            "Role": role,
            "Qualification": qual
        }

        # Check preferences 1 to 5
        for rank in range(1, 6):
            pref_val = None
            exp_val = 0.0

            for col in df.columns:
                col_clean = str(col).lower().replace("_", " ").strip()
                # Check for preference course name (e.g. Course Preference 1, Preference 1)
                is_exp_col = any(k in col_clean for k in ["exp", "year", "teaching", "duration"])
                if not is_exp_col and (f"preference {rank}" in col_clean or f"pref {rank}" in col_clean or col_clean == f"course {rank}"):
                    if pd.notna(row[col]) and str(row[col]).strip().lower() not in ('nan', 'none', ''):
                        pref_val = str(row[col]).strip()
                
                # Check for course experience (e.g. Experience in Course 1, Course 1 Experience, Exp 1)
                if is_exp_col and (f"{rank}" in col_clean or f"course {rank}" in col_clean or f"pref {rank}" in col_clean):
                    val = pd.to_numeric(row[col], errors='coerce')
                    if pd.notna(val):
                        exp_val = float(val)

            p_dict[f"Preference_{rank}"] = pref_val

            if pref_val:
                exp_rows.append({
                    "Faculty_ID": fac_id,
                    "Subject": pref_val,
                    "Qualification": qual,
                    "Skill_Level": derive_skill_level(exp_val),
                    "Experience_Years": exp_val,
                    "Preference_Rank": rank,
                    "Role": role
                })

        pref_rows.append(p_dict)

    pref_df = pd.DataFrame(pref_rows)
    exp_df = pd.DataFrame(exp_rows)

    return pref_df, exp_df

def read_faculty_preferences(file_source: Union[str, BinaryIO]) -> pd.DataFrame:
    """Reads faculty preferences Excel file into a cleaned DataFrame."""
    df = pd.read_excel(file_source)
    df.columns = [str(col).strip() for col in df.columns]

    if any("course preference" in c.lower() for c in df.columns):
        pref_df, _ = parse_wide_combined_sheet(df)
        return pref_df

    expected_cols = ["Faculty_ID", "Faculty_Name", "Preference_1", "Preference_2", "Preference_3", "Preference_4", "Preference_5"]
    for col in expected_cols:
        if col not in df.columns:
            for actual_col in df.columns:
                if actual_col.lower() == col.lower():
                    df.rename(columns={actual_col: col}, inplace=True)
                    break

    for col in df.columns:
        if df[col].dtype == object or df[col].dtype == 'string':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'nan': None, 'None': None, '': None})

    df.dropna(subset=["Faculty_ID"], inplace=True)
    if "Role" not in df.columns:
        df["Role"] = df.apply(lambda r: extract_role(str(r.get("Designation", "")), str(r.get("Additional_Duties", "")), str(r.get("Faculty_Name", ""))), axis=1)

    return df

def read_faculty_experience(file_source: Union[str, BinaryIO]) -> pd.DataFrame:
    """Reads faculty experience & qualification Excel file into a cleaned DataFrame."""
    df = pd.read_excel(file_source)
    df.columns = [str(col).strip() for col in df.columns]

    if any("course preference" in c.lower() for c in df.columns):
        _, exp_df = parse_wide_combined_sheet(df)
        return exp_df

    expected_cols = ["Faculty_ID", "Subject", "Qualification", "Skill_Level", "Experience_Years"]
    for col in expected_cols:
        if col not in df.columns:
            for actual_col in df.columns:
                if actual_col.lower() == col.lower():
                    df.rename(columns={actual_col: col}, inplace=True)
                    break

    for col in ["Faculty_ID", "Subject", "Qualification", "Skill_Level"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'nan': None, 'None': None, '': None})

    if "Experience_Years" in df.columns:
        df["Experience_Years"] = pd.to_numeric(df["Experience_Years"], errors='coerce').fillna(0)

    df.dropna(subset=["Faculty_ID", "Subject"], inplace=True)
    return df

def load_both_excel_files(pref_file: Union[str, BinaryIO], exp_file: Union[str, BinaryIO]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Loads and returns cleaned preferences and experience DataFrames."""
    if isinstance(pref_file, str):
        df_check = pd.read_excel(pref_file)
        if any("course preference" in c.lower() for c in df_check.columns):
            return parse_wide_combined_sheet(df_check)

    pref_df = read_faculty_preferences(pref_file)
    exp_df = read_faculty_experience(exp_file)
    return pref_df, exp_df

def load_all_excel_files(
    pref_file: Union[str, BinaryIO],
    exp_file: Union[str, BinaryIO],
    course_file: Optional[Union[str, BinaryIO]] = None
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Loads preferences, experience, and course details DataFrames."""
    pref_df, exp_df = load_both_excel_files(pref_file, exp_file)
    if course_file is not None:
        course_df = read_course_details(course_file)
    else:
        import os
        default_course_path = "input/course_details.xlsx"
        if os.path.exists(default_course_path):
            course_df = read_course_details(default_course_path)
        else:
            records = [
                {"Subject": "Quantitative Aptitude & Logical Reasoning", "L": 1, "T": 2, "P": 0, "C": 2, "L-T-P-C": "1-2-0-2", "Total_Hours": 3},
                {"Subject": "Deep Learning", "L": 3, "T": 0, "P": 2, "C": 4, "L-T-P-C": "3-0-2-4", "Total_Hours": 5},
                {"Subject": "Web Technologies", "L": 2, "T": 0, "P": 4, "C": 4, "L-T-P-C": "2-0-4-4", "Total_Hours": 6},
                {"Subject": "Inter Departmental Project with Agentic Tools", "L": 0, "T": 0, "P": 2, "C": 1, "L-T-P-C": "0-0-2-1", "Total_Hours": 2},
                {"Subject": "Machine Learning Operations (DE-2)", "L": 2, "T": 0, "P": 2, "C": 3, "L-T-P-C": "2-0-2-3", "Total_Hours": 4},
                {"Subject": "Computer Vision (DE-2)", "L": 3, "T": 0, "P": 2, "C": 4, "L-T-P-C": "3-0-2-4", "Total_Hours": 5}
            ]
            course_df = pd.DataFrame(records)

    return pref_df, exp_df, course_df

