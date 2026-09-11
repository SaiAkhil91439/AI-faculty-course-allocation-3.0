import os
import pandas as pd

def generate_sample_files(output_dir: str = "input"):
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Course Details Dataset with L-T-P-C structure matching user's specification
    course_sample_data = [
        {"Courses": "Quantitative Aptitude & Logical Reasoning", "L-T-P-C": "1-2-0-2"},
        {"Courses": "Deep Learning", "L-T-P-C": "3-0-2-4"},
        {"Courses": "Web Technologies", "L-T-P-C": "2-0-4-4"},
        {"Courses": "Inter Departmental Project with Agentic Tools", "L-T-P-C": "0-0-2-1"},
        {"Courses": "Machine Learning Operations (DE-2)", "L-T-P-C": "2-0-2-3"},
        {"Courses": "Computer Vision (DE-2)", "L-T-P-C": "3-0-2-4"}
    ]
    course_df = pd.DataFrame(course_sample_data)
    course_path = os.path.join(output_dir, "course_details.xlsx")
    course_df.to_excel(course_path, index=False)

    # 2. Combined Wide Format Excel Dataset matching user's input specification
    wide_sample_data = [
        {
            "Faculty Name": "Dr. M. Venkateswarlu",
            "Designation": "Professor & Dean",
            "Additional Duties": "Dean",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 12,
            "Course Preference 2": "Computer Vision (DE-2)", "Experience in Course 2": 10,
            "Course Preference 3": "Web Technologies", "Experience in Course 3": 8,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 5,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 3
        },
        {
            "Faculty Name": "Dr. M. Sreenivasa Rao",
            "Designation": "Professor & HOD",
            "Additional Duties": "HOD",
            "Qualification": "PhD AI",
            "Course Preference 1": "Computer Vision (DE-2)", "Experience in Course 1": 11,
            "Course Preference 2": "Deep Learning", "Experience in Course 2": 9,
            "Course Preference 3": "Machine Learning Operations (DE-2)", "Experience in Course 3": 7,
            "Course Preference 4": "Web Technologies", "Experience in Course 4": 6,
            "Course Preference 5": "Inter Departmental Project with Agentic Tools", "Experience in Course 5": 4
        },
        {
            "Faculty Name": "Dr. D. Arban Das",
            "Designation": "Professor & Vice Dean",
            "Additional Duties": "Vice Dean",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Web Technologies", "Experience in Course 1": 10,
            "Course Preference 2": "Machine Learning Operations (DE-2)", "Experience in Course 2": 8,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 6,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 4,
            "Course Preference 5": "Computer Vision (DE-2)", "Experience in Course 5": 2
        },
        {
            "Faculty Name": "Dr. Siva",
            "Designation": "Professor",
            "Additional Duties": "None",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 10,
            "Course Preference 2": "Computer Vision (DE-2)", "Experience in Course 2": 7,
            "Course Preference 3": "Web Technologies", "Experience in Course 3": 5,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 2
        },
        {
            "Faculty Name": "Dr. Priya Verma",
            "Designation": "Professor",
            "Additional Duties": "None",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 9,
            "Course Preference 2": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 2": 8,
            "Course Preference 3": "Computer Vision (DE-2)", "Experience in Course 3": 6,
            "Course Preference 4": "Web Technologies", "Experience in Course 4": 4,
            "Course Preference 5": "Machine Learning Operations (DE-2)", "Experience in Course 5": 2
        },
        {
            "Faculty Name": "Prof. Kumar Swamy",
            "Designation": "Professor",
            "Additional Duties": "None",
            "Qualification": "PhD DBMS",
            "Course Preference 1": "Web Technologies", "Experience in Course 1": 10,
            "Course Preference 2": "Machine Learning Operations (DE-2)", "Experience in Course 2": 7,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 5,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 3,
            "Course Preference 5": "Deep Learning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Anish Reddy",
            "Designation": "Associate Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech CSE",
            "Course Preference 1": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 1": 7,
            "Course Preference 2": "Machine Learning Operations (DE-2)", "Experience in Course 2": 6,
            "Course Preference 3": "Deep Learning", "Experience in Course 3": 4,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 3,
            "Course Preference 5": "Web Technologies", "Experience in Course 5": 2
        },
        {
            "Faculty Name": "Prof. Sunita Gupta",
            "Designation": "Associate Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech Cloud",
            "Course Preference 1": "Machine Learning Operations (DE-2)", "Experience in Course 1": 8,
            "Course Preference 2": "Web Technologies", "Experience in Course 2": 6,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 5,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 2,
            "Course Preference 5": "Computer Vision (DE-2)", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Rajesh Iyer",
            "Designation": "Associate Professor",
            "Additional Duties": "None",
            "Qualification": "PhD Data Science",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 7,
            "Course Preference 2": "Computer Vision (DE-2)", "Experience in Course 2": 6,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 5,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 3,
            "Course Preference 5": "Web Technologies", "Experience in Course 5": 2
        },
        {
            "Faculty Name": "Prof. Meena Patel",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech CSE",
            "Course Preference 1": "Web Technologies", "Experience in Course 1": 5,
            "Course Preference 2": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 2": 4,
            "Course Preference 3": "Machine Learning Operations (DE-2)", "Experience in Course 3": 3,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 2,
            "Course Preference 5": "Computer Vision (DE-2)", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Vikram Singh",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD IT",
            "Course Preference 1": "Machine Learning Operations (DE-2)", "Experience in Course 1": 6,
            "Course Preference 2": "Deep Learning", "Experience in Course 2": 4,
            "Course Preference 3": "Computer Vision (DE-2)", "Experience in Course 3": 3,
            "Course Preference 4": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 4": 2,
            "Course Preference 5": "Web Technologies", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Lakshmi N",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD AI",
            "Course Preference 1": "Computer Vision (DE-2)", "Experience in Course 1": 6,
            "Course Preference 2": "Web Technologies", "Experience in Course 2": 4,
            "Course Preference 3": "Deep Learning", "Experience in Course 3": 3,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Prof. Suresh Kumar",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech CSE",
            "Course Preference 1": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 1": 5,
            "Course Preference 2": "Web Technologies", "Experience in Course 2": 4,
            "Course Preference 3": "Machine Learning Operations (DE-2)", "Experience in Course 3": 3,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Deep Learning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Kavita Sharma",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 6,
            "Course Preference 2": "Computer Vision (DE-2)", "Experience in Course 2": 5,
            "Course Preference 3": "Web Technologies", "Experience in Course 3": 3,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Prof. Ramesh Babu",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech AI",
            "Course Preference 1": "Computer Vision (DE-2)", "Experience in Course 1": 5,
            "Course Preference 2": "Deep Learning", "Experience in Course 2": 4,
            "Course Preference 3": "Machine Learning Operations (DE-2)", "Experience in Course 3": 3,
            "Course Preference 4": "Web Technologies", "Experience in Course 4": 2,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Neha Kulkarni",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD Data Science",
            "Course Preference 1": "Web Technologies", "Experience in Course 1": 6,
            "Course Preference 2": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 2": 5,
            "Course Preference 3": "Deep Learning", "Experience in Course 3": 3,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Machine Learning Operations (DE-2)", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Prof. Amit Shah",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech Cloud",
            "Course Preference 1": "Machine Learning Operations (DE-2)", "Experience in Course 1": 5,
            "Course Preference 2": "Web Technologies", "Experience in Course 2": 4,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 3,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 2,
            "Course Preference 5": "Computer Vision (DE-2)", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Geeta Nair",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD IT",
            "Course Preference 1": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 1": 6,
            "Course Preference 2": "Machine Learning Operations (DE-2)", "Experience in Course 2": 4,
            "Course Preference 3": "Computer Vision (DE-2)", "Experience in Course 3": 3,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 2,
            "Course Preference 5": "Web Technologies", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Prof. Nitin Desai",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech CSE",
            "Course Preference 1": "Deep Learning", "Experience in Course 1": 4,
            "Course Preference 2": "Web Technologies", "Experience in Course 2": 3,
            "Course Preference 3": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 3": 2,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Machine Learning Operations (DE-2)", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Sunita Rao",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD CSE",
            "Course Preference 1": "Computer Vision (DE-2)", "Experience in Course 1": 5,
            "Course Preference 2": "Deep Learning", "Experience in Course 2": 4,
            "Course Preference 3": "Web Technologies", "Experience in Course 3": 3,
            "Course Preference 4": "Machine Learning Operations (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Dr. Pooja Sharma",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "PhD AI",
            "Course Preference 1": "Inter Departmental Project with Agentic Tools", "Experience in Course 1": 6,
            "Course Preference 2": "Machine Learning Operations (DE-2)", "Experience in Course 2": 5,
            "Course Preference 3": "Deep Learning", "Experience in Course 3": 4,
            "Course Preference 4": "Computer Vision (DE-2)", "Experience in Course 4": 2,
            "Course Preference 5": "Web Technologies", "Experience in Course 5": 1
        },
        {
            "Faculty Name": "Prof. Alok Kumar",
            "Designation": "Assistant Professor",
            "Additional Duties": "None",
            "Qualification": "M.Tech CSE",
            "Course Preference 1": "Inter Departmental Project with Agentic Tools", "Experience in Course 1": 5,
            "Course Preference 2": "Quantitative Aptitude & Logical Reasoning", "Experience in Course 2": 4,
            "Course Preference 3": "Web Technologies", "Experience in Course 3": 3,
            "Course Preference 4": "Deep Learning", "Experience in Course 4": 2,
            "Course Preference 5": "Computer Vision (DE-2)", "Experience in Course 5": 1
        }
    ]

    wide_df = pd.DataFrame(wide_sample_data)
    wide_path = os.path.join(output_dir, "faculty_sample_combined.xlsx")
    pref_path = os.path.join(output_dir, "faculty_preferences.xlsx")
    exp_path = os.path.join(output_dir, "faculty_experience.xlsx")

    # Save wide format sample
    wide_df.to_excel(wide_path, index=False)
    
    # Save standard preferences & experience breakdown for backward compatibility
    pref_rows = []
    exp_rows = []
    for idx, row in wide_df.iterrows():
        fac_id = f"F{idx+1:03d}"
        pref_item = {
            "Faculty_ID": fac_id,
            "Faculty_Name": row["Faculty Name"],
            "Designation": row["Designation"],
            "Additional_Duties": row.get("Additional Duties", "None"),
            "Qualification": row["Qualification"]
        }
        for r in range(1, 6):
            pref_item[f"Preference_{r}"] = row[f"Course Preference {r}"]
            exp_rows.append({
                "Faculty_ID": fac_id,
                "Subject": row[f"Course Preference {r}"],
                "Qualification": row["Qualification"],
                "Skill_Level": "Expert" if row[f"Experience in Course {r}"] >= 8 else ("Advanced" if row[f"Experience in Course {r}"] >= 5 else "Intermediate"),
                "Experience_Years": row[f"Experience in Course {r}"]
            })
        pref_rows.append(pref_item)

    pd.DataFrame(pref_rows).to_excel(pref_path, index=False)
    pd.DataFrame(exp_rows).to_excel(exp_path, index=False)
    print(f"Sample files generated at {course_path} and {wide_path}")

if __name__ == "__main__":
    generate_sample_files()

