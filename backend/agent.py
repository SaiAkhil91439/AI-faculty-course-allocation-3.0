import os
import json
from typing import Dict, Any, Optional
import pandas as pd
from backend.excel_reader import load_all_excel_files
from backend.validator import validate_inputs
from backend.scoring import calculate_compatibility_scores
from backend.optimizer import solve_faculty_allocation
from backend.excel_generator import generate_allocation_excel

try:
    from google import genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

class FacultyAllocationAgent:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        self.client = None
        self.latest_allocation_res = None
        self.latest_scores_df = None
        self.scenario_history = []
        if HAS_GENAI and self.api_key:
            try:
                self.client = genai.Client(api_key=self.api_key)
            except Exception as e:
                print(f"Gemini client initialization warning: {e}")

    def run_pipeline(
        self,
        pref_file_source: Any,
        exp_file_source: Optional[Any] = None,
        num_sections: int = 12,
        num_subjects: int = 6,
        sections_per_faculty: int = 3,
        max_subjects_per_faculty: int = 1,
        course_file_source: Optional[Any] = None,
        dean_max_sections: int = 2,
        hod_max_sections: int = 3,
        high_exp_threshold: float = 8.0,
        weights: Optional[Dict[str, float]] = None,
        output_excel_path: str = "output/faculty_allocation.xlsx"
    ) -> Dict[str, Any]:
        """
        Full Agent Orchestration Workflow:
        1. Read Excel files & Course L-T-P-C Structure
        2. Validate Data & Designation Hours Constraints
        3. Score Compatibility (Strict 1st preference priority & least-preferred subject indexing)
        4. Optimize with OR-Tools CP-SAT (Designation bounds, High-Experience 2+1 split & 1 L + 2 T/P per section)
        5. Generate Output Excel
        6. Synthesize AI Explanation
        """
        agent_steps = []

        # Step 1: Read Files
        agent_steps.append({"step": 1, "action": "Read Excel Files, Course L-T-P-C & Extract Roles", "status": "In Progress"})
        pref_df, exp_df, course_df = load_all_excel_files(pref_file_source, exp_file_source, course_file_source)
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = f"Loaded {len(pref_df)} faculty preferences, {len(exp_df)} experience records, and {len(course_df)} course L-T-P-C details."

        # Step 2: Validate Data
        agent_steps.append({"step": 2, "action": "Validate Data & Designation Workload Limits", "status": "In Progress"})
        val_info = validate_inputs(
            pref_df=pref_df,
            exp_df=exp_df,
            num_sections=num_sections,
            num_subjects=num_subjects,
            sections_per_faculty=sections_per_faculty,
            max_subjects_per_faculty=max_subjects_per_faculty,
            course_df=course_df,
            dean_max_sections=dean_max_sections,
            hod_max_sections=hod_max_sections,
            high_exp_threshold=high_exp_threshold
        )
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = (
            f"Validation passed! Identified {val_info['deans_count']} Dean(s), {val_info['hods_count']} HOD(s), "
            f"and {val_info['high_exp_count']} High-Experience Faculty (course teaching exp >= {high_exp_threshold} yrs)."
        )

        # Step 3: Calculate Scores
        agent_steps.append({"step": 3, "action": "Calculate Compatibility & Least-Preferred Subject Index", "status": "In Progress"})
        scores_res = calculate_compatibility_scores(pref_df, exp_df, weights=weights)
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = (
            f"Computed compatibility scores for {len(scores_res['scores_df'])} faculty-subject pairs. "
            f"Prioritized 1st preference ranks and indexed least-preferred subjects for experienced faculty allocation."
        )

        # Step 4: Call Optimization Engine
        agent_steps.append({"step": 4, "action": "Solve Allocation with Google OR-Tools CP-SAT", "status": "In Progress"})
        allocation_res = solve_faculty_allocation(
            pref_df=pref_df,
            exp_df=exp_df,
            scores_dict=scores_res,
            num_sections=num_sections,
            num_subjects=num_subjects,
            sections_per_faculty=sections_per_faculty,
            max_subjects_per_faculty=max_subjects_per_faculty,
            course_df=course_df,
            dean_max_sections=dean_max_sections,
            hod_max_sections=hod_max_sections,
            high_exp_threshold=high_exp_threshold
        )
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = (
            f"OR-Tools solver status: {allocation_res['status']}. "
            f"Average score: {allocation_res['average_score']}/100. Applied 2+1 section split rule to {allocation_res.get('high_exp_faculty_count', 0)} high-experience faculty."
        )

        # Step 5: Generate Excel
        agent_steps.append({"step": 5, "action": "Generate Workload Excel Workbook", "status": "In Progress"})
        excel_path = generate_allocation_excel(allocation_res, output_excel_path)
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = f"Excel report matching target workload table saved to {excel_path}."

        # Step 6: AI Explanation
        agent_steps.append({"step": 6, "action": "Generate Allocation Executive Summary", "status": "In Progress"})
        ai_explanation = self._generate_explanation(allocation_res, scores_res["scores_df"])
        agent_steps[-1]["status"] = "Completed"
        agent_steps[-1]["details"] = "Executive summary synthesized successfully."

        # Save active run state for live chatbot & scenario comparisons
        self.latest_allocation_res = allocation_res
        self.latest_scores_df = scores_res["scores_df"]

        scenario_entry = {
            "id": f"Scenario {len(self.scenario_history) + 1}",
            "weights": weights or {"preference": 0.4, "experience": 0.3, "skill": 0.2, "qualification": 0.1},
            "num_sections": allocation_res["num_sections"],
            "num_subjects": allocation_res["num_subjects"],
            "average_score": allocation_res["average_score"],
            "total_assignments": allocation_res["total_assignments"],
            "high_exp_faculty_count": allocation_res.get("high_exp_faculty_count", 0),
            "status": allocation_res["status"]
        }
        self.scenario_history.append(scenario_entry)

        return {
            "success": True,
            "agent_steps": agent_steps,
            "summary": {
                "num_sections": allocation_res["num_sections"],
                "num_subjects": allocation_res["num_subjects"],
                "sections_per_faculty": allocation_res["sections_per_faculty"],
                "faculty_required_per_subject": allocation_res["faculty_required_per_subject"],
                "high_exp_threshold": allocation_res.get("high_exp_threshold", high_exp_threshold),
                "high_exp_faculty_count": allocation_res.get("high_exp_faculty_count", 0),
                "total_assignments": allocation_res["total_assignments"],
                "average_score": allocation_res["average_score"],
                "status": allocation_res["status"],
                "excel_file": excel_path,
                "scenario_id": scenario_entry["id"]
            },
            "allocation_data": allocation_res,
            "ai_explanation": ai_explanation
        }

    def _generate_explanation(self, allocation_res: Dict[str, Any], scores_df: pd.DataFrame) -> str:
        """Generates natural language executive summary explaining L-T-P-C, high-experience 2+1 split, and designation hour bounds."""
        fac_allocs = allocation_res["faculty_allocations"]
        avg_score = allocation_res["average_score"]
        subjects = allocation_res["subjects"]
        c_details = allocation_res.get("course_details", {})
        high_exp_count = allocation_res.get("high_exp_faculty_count", 0)

        summary_text = (
            f"### Executive Allocation Summary\n"
            f"- **Optimization Status**: {allocation_res['status']}\n"
            f"- **Average Compatibility Score**: **{avg_score} / 100**\n"
            f"- **Sections**: {allocation_res['num_sections']} | **Subjects**: {allocation_res['num_subjects']}\n"
            f"- **High-Experience Rule Applied**: **{high_exp_count} faculty** with course teaching exp >= {allocation_res.get('high_exp_threshold', 8.0)} yrs received **2 sections in their highest-experienced course + 1 section in their 2nd highest-experienced course**.\n"
            f"- **Designation Workload Hours Applied**:\n"
            f"  - **Deans**: 4–6 hrs/wk\n"
            f"  - **HODs**: 10–12 hrs/wk\n"
            f"  - **Associate Professors**: 14–16 hrs/wk\n"
            f"  - **Assistant Professors**: 16–18 hrs/wk\n"
            f"- **Component Ratio**: Each section has 1 Main Lecture (L) faculty and 2 Tutorial/Practical (T/P) faculty.\n\n"
            f"### Subject L-T-P-C & Assignment Breakdown\n"
        )

        highlights = []
        for subj in subjects:
            ltpc = c_details.get(subj, {}).get("L-T-P-C", "N/A")
            tot_h = c_details.get(subj, {}).get("Total_Hours", 5)
            subj_allocs = [a for a in fac_allocs if a["Subject"] == subj]
            assigned_str = ", ".join([
                f"**{a['Lecture_Faculty_Name']}** ({a['Lecture_Designation']} - {a['Sections']} [{a['Num_Sections']} Secs, {a['Experience_Years']} yrs exp] - {a['Assigned_Weekly_Hours']} hrs/wk)"
                for a in subj_allocs
            ])
            highlights.append(f"- **{subj}** `[{ltpc}]` ({tot_h} hrs/sec): Assigned -> {assigned_str}")

        details_body = "\n".join(highlights)
        fallback_explanation = summary_text + details_body

        if self.client:
            try:
                prompt = (
                    "You are an AI Faculty Allocation Agent. Analyze the following optimal allocation result "
                    "produced by Google OR-Tools and write a professional executive summary explaining how L-T-P-C course structures "
                    "and designation workload hours (Deans: 4-6 hrs, HODs: 10-12 hrs, Associate Profs: 14-16 hrs, Assistant Profs: 16-18 hrs) "
                    "were satisfied.\n\n"
                    f"Data JSON:\n{json.dumps(fac_allocs, indent=2)}\n\n"
                    "Provide a clear, formatted markdown summary."
                )
                response = self.client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )
                if response and response.text:
                    return response.text
            except Exception as e:
                print(f"Gemini generation call failed, using fallback: {e}")

        return fallback_explanation

    def handle_chat_query(self, user_question: str) -> Dict[str, Any]:
        """
        Processes natural language questions from Deans & HODs about active allocation results,
        workload distribution, non-assignments, and scenario comparisons.
        """
        if not self.latest_allocation_res:
            return {
                "success": False,
                "answer": "⚠️ No allocation pipeline has been executed yet. Please click **'Run Full Allocation Agent'** first to generate allocation data.",
                "suggested_questions": ["Run allocation pipeline first"]
            }

        alloc_data = self.latest_allocation_res
        fac_allocs = alloc_data.get("faculty_allocations", [])
        scenarios = self.scenario_history
        q_lower = user_question.lower().strip()

        # 1. Ask Gemini LLM if client is available
        if self.client:
            try:
                prompt = (
                    "You are the Conversational AI Assistant for an Academic Faculty-Subject-Section Allocation Agent.\n"
                    "Help the Dean/HOD by answering their question accurately based on the allocation data provided below.\n\n"
                    f"ACTIVE ALLOCATION METRICS:\n"
                    f"- Status: {alloc_data['status']}\n"
                    f"- Average Compatibility Score: {alloc_data['average_score']}/100\n"
                    f"- Sections: {alloc_data['num_sections']} | Subjects: {alloc_data['num_subjects']}\n"
                    f"- High Exp Faculty Count (>= {alloc_data.get('high_exp_threshold', 8.0)} yrs): {alloc_data.get('high_exp_faculty_count', 0)}\n\n"
                    f"FACULTY ASSIGNMENTS LIST:\n"
                    f"{json.dumps(fac_allocs, indent=2)}\n\n"
                    f"SCENARIO HISTORY FOR A/B TESTING:\n"
                    f"{json.dumps(scenarios, indent=2)}\n\n"
                    f"USER QUESTION: {user_question}\n\n"
                    "Respond concisely with professional markdown formatting (use bullet points or tables where appropriate)."
                )
                response = self.client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )
                if response and response.text:
                    return {
                        "success": True,
                        "answer": response.text,
                        "suggested_questions": [
                            "Show faculty teaching > 15 hours/week",
                            "Why wasn't Dr. Smith assigned to Machine Learning?",
                            "Compare current run with previous scenario"
                        ]
                    }
            except Exception as e:
                print(f"Gemini chat generation failed: {e}")

        # 2. Rule-based Analytical Fallback Engine
        # Query Type A: Teaching Hours / Workload query (e.g. > 15 hours, teaching load)
        if any(kw in q_lower for kw in ["hour", "workload", "load", "teaching", "15", "heavy"]):
            high_hours = [a for a in fac_allocs if a.get("Assigned_Weekly_Hours", 0) >= 14]
            if not high_hours:
                answer = "📊 All faculty have balanced workload within normal limits (< 14 hrs/week)."
            else:
                lines = [
                    "### 📊 Faculty Workload Breakdown (>= 14 Hours/Week)",
                    "| Faculty Name | Designation | Subject | Assigned Sections | Weekly Hours |",
                    "| :--- | :--- | :--- | :--- | :--- |"
                ]
                for a in high_hours:
                    lines.append(f"| **{a['Lecture_Faculty_Name']}** | {a['Lecture_Designation']} | {a['Subject']} | {a['Sections']} | **{a['Assigned_Weekly_Hours']} hrs/wk** |")
                lines.append("\n> **Note:** Designation caps enforced: Assistant Profs (16-18 hrs), Associate Profs (14-16 hrs), HODs (10-12 hrs), Deans (4-6 hrs).")
                answer = "\n".join(lines)
            
            return {
                "success": True,
                "answer": answer,
                "suggested_questions": [
                    "Show HOD and Dean workload",
                    "Compare current run with previous scenario",
                    "List all subject assignments"
                ]
            }

        # Query Type B: Scenario Comparison / A/B Testing
        elif any(kw in q_lower for kw in ["compare", "scenario", "previous", "a/b", "weight", "run"]):
            if len(scenarios) <= 1:
                answer = (
                    "### 🔄 Scenario Comparison (A/B Testing)\n"
                    "You currently have **1 allocation run** recorded.\n\n"
                    "💡 **Tip:** Adjust scoring weight sliders in the top dashboard and click **'Run Full Allocation Agent'** again to generate Scenario 2 for side-by-side comparison!"
                )
            else:
                lines = [
                    "### 🔄 Scenario Comparison (A/B Testing Results)",
                    "| Scenario ID | Weights (Pref/Exp/Skill/Qual) | Avg Score | Total Assignments | High Exp Faculty | Status |",
                    "| :--- | :--- | :--- | :--- | :--- | :--- |"
                ]
                for s in scenarios:
                    w = s.get("weights", {})
                    w_str = f"P:{int(w.get('preference',0)*100)}% E:{int(w.get('experience',0)*100)}% S:{int(w.get('skill',0)*100)}% Q:{int(w.get('qualification',0)*100)}%"
                    lines.append(f"| **{s['id']}** | {w_str} | **{s['average_score']}/100** | {s['total_assignments']} | {s['high_exp_faculty_count']} | `{s['status']}` |")
                
                s_first = scenarios[0]
                s_last = scenarios[-1]
                diff = round(s_last['average_score'] - s_first['average_score'], 2)
                diff_str = f"+{diff}" if diff > 0 else f"{diff}"
                lines.append(f"\n📈 **Score Impact:** From {s_first['id']} to {s_last['id']}, average compatibility score changed by **{diff_str} points**.")
                answer = "\n".join(lines)

            return {
                "success": True,
                "answer": answer,
                "suggested_questions": [
                    "Show faculty teaching > 15 hours/week",
                    "Why wasn't Dr. Smith assigned to Machine Learning?",
                    "List all subject assignments"
                ]
            }

        # Query Type C: Why / Non-assignment query
        elif any(kw in q_lower for kw in ["why", "not assigned", "assign", "smith", "reason"]):
            answer = (
                "### 🔍 Assignment Rationale Analysis\n"
                "The OR-Tools solver maximizes the global compatibility score subject to strict academic constraints:\n\n"
                "1. **Preference Priority:** 1st preference ranks carry 100 points, while 5th preference carries only 10 points.\n"
                "2. **Experience & Qualification Filtering:** Non-qualified faculty are strictly excluded.\n"
                "3. **High-Experience 2+1 Rule:** Faculty with >= 8 years experience get 2 sections in top-1 course + 1 section in top-2 course.\n"
                "4. **Workload Caps:** Deans (max 2 sections), HODs (max 3 sections), preventing over-allocation.\n\n"
                "To inspect exact scores for any faculty-subject pair, download the **`faculty_allocation.xlsx`** workbook and view the **Detailed Breakdown** sheet."
            )
            return {
                "success": True,
                "answer": answer,
                "suggested_questions": [
                    "Show faculty teaching > 15 hours/week",
                    "Compare current run with previous scenario",
                    "Summarize allocation status"
                ]
            }

        # Default fallback summary
        else:
            lines = [
                f"### 🤖 Allocation AI Assistant Overview",
                f"- **Optimization Status**: `{alloc_data['status']}`",
                f"- **Average Compatibility Score**: **{alloc_data['average_score']} / 100**",
                f"- **Total Assignments**: {alloc_data['total_assignments']}",
                f"- **High-Experience Faculty Handled**: {alloc_data.get('high_exp_faculty_count', 0)}",
                f"\nFeel free to ask me questions like:",
                f"- *\"Show faculty teaching > 15 hours/week\"*",
                f"- *\"Compare current run with previous scenario\"*",
                f"- *\"Why wasn't a specific faculty assigned to a subject?\"*"
            ]
            return {
                "success": True,
                "answer": "\n".join(lines),
                "suggested_questions": [
                    "Show faculty teaching > 15 hours/week",
                    "Compare current run with previous scenario",
                    "Why wasn't Dr. Smith assigned to Machine Learning?"
                ]
            }

