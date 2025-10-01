import os
import json
import pandas as pd
from fpdf import FPDF
import matplotlib.pyplot as plt
from typing import Dict, Any, List, Union


class OutputWriter:
    """
    Utility class to save agent outputs into Excel, JSON, PDF, or charts.
    Handles formatting differences for Meeting Prep, Deal Execution, Coaching.
    """

    def __init__(self, base_dir: str = "outputs"):
        self.base_dir = base_dir
        self.ensure_directories()

    def ensure_directories(self):
        folders = [
            self.base_dir,
            os.path.join(self.base_dir, "prep_packs"),
            os.path.join(self.base_dir, "followups"),
            os.path.join(self.base_dir, "charts"),
            os.path.join(self.base_dir, "deal_summaries"),
            os.path.join(self.base_dir, "coaching_reports"),
        ]
        for folder in folders:
            os.makedirs(folder, exist_ok=True)

    # -------- Excel Helpers --------
    def append_to_excel(self, data: Dict[str, Any], file_name: str, sheet_name: str = "Sheet1"):
        path = os.path.join(self.base_dir, file_name)
        new_df = pd.DataFrame([data])

        if os.path.exists(path):
            with pd.ExcelWriter(path, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                try:
                    existing_df = pd.read_excel(path, sheet_name=sheet_name)
                    combined = pd.concat([existing_df, new_df], ignore_index=True)
                    combined.to_excel(writer, sheet_name=sheet_name, index=False)
                except ValueError:
                    new_df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            new_df.to_excel(path, sheet_name=sheet_name, index=False)
        return path

    # -------- PDF Helpers --------
    def render_table(self, pdf: FPDF, data: Dict[str, Union[int, float, str]], col_widths=(80, 40)):
        pdf.set_font("DejaVu", size=12)
        pdf.set_fill_color(200, 200, 200)  # header bg
        pdf.cell(col_widths[0], 10, "Key", border=1, fill=True, align="C")
        pdf.cell(col_widths[1], 10, "Value", border=1, fill=True, align="C")
        pdf.ln()

        pdf.set_fill_color(245, 245, 245)
        fill = False
        for k, v in data.items():
            pdf.cell(col_widths[0], 10, str(k), border=1, fill=fill)
            pdf.cell(col_widths[1], 10, str(v), border=1, fill=fill)
            pdf.ln()
            fill = not fill

    def format_and_write(self, pdf: FPDF, content: Any, title: str = None):
        if title:
            pdf.set_font("DejaVu", "B", 12)
            pdf.cell(0, 10, title, ln=True)

        if isinstance(content, list):
            pdf.set_font("DejaVu", size=12)
            for item in content:
                pdf.multi_cell(0, 8, "• " + str(item))
        elif isinstance(content, dict):
            if all(isinstance(v, (int, float, str)) for v in content.values()):
                self.render_table(pdf, content)
            else:
                for k, v in content.items():
                    self.format_and_write(pdf, v, title=str(k))
        else:
            pdf.set_font("DejaVu", size=12)
            pdf.multi_cell(0, 8, str(content))

    def save_radar_chart(self, scores: Dict[str, int], title: str, file_name: str) -> str:
        labels = list(scores.keys())
        values = list(scores.values())
        values += values[:1]
        angles = [n / float(len(labels)) * 2 * 3.14159 for n in range(len(labels))]
        angles += angles[:1]

        fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
        ax.plot(angles, values, linewidth=2)
        ax.fill(angles, values, alpha=0.25)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(labels)
        ax.set_title(title)

        path = os.path.join(self.base_dir, "charts", file_name)
        plt.savefig(path)
        plt.close(fig)
        return path

    def write_to_pdf(self, data: Any, file_name: str, subfolder: str = "", meta: Dict[str, str] = None):
        folder = os.path.join(self.base_dir, subfolder) if subfolder else self.base_dir
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, file_name)

        pdf = FPDF()
        pdf.add_page()

        # Add Unicode font
        pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
        pdf.add_font("DejaVu", "B", "DejaVuSans-Bold.ttf", uni=True)
        pdf.set_font("DejaVu", size=12)

        if meta:
            pdf.set_font("DejaVu", "B", 14)
            header = f"{meta.get('client','')} <> Dice || {meta.get('report_type','Report')}"
            pdf.cell(0, 10, header, ln=True)
            pdf.set_font("DejaVu", size=12)
            pdf.cell(0, 8, f"Meeting ID: {meta.get('meeting_id','')}", ln=True)
            if meta.get("date"):
                pdf.cell(0, 8, f"Date: {meta.get('date')}", ln=True)
            pdf.ln(5)

        if isinstance(data, str):
            for line in data.split("\n"):
                pdf.multi_cell(0, 8, line)
        # elif isinstance(data, dict):
        #     for k, v in data.items():
        #         self.format_and_write(pdf, v, title=str(k))
                
        elif isinstance(data, dict):
            # --- Special formatting for Meeting Preparation Pack ---
            if "prep_pack" in data:  
                pdf.set_font("DejaVu", "B", 16)
                pdf.cell(0, 10, "Meeting Preparation Pack", ln=True)

                self.format_and_write(pdf, data.get("summary_till_now", ""), title="Summary Till Now")
                self.format_and_write(pdf, data.get("next_steps", []), title="Next Steps")

                prep = data.get("prep_pack", {})
                self.format_and_write(pdf, prep.get("agenda", []), title="Agenda")
                self.format_and_write(pdf, prep.get("key_stakeholders", {}), title="Key Stakeholders")
                self.format_and_write(pdf, prep.get("objections", []), title="Likely Objections")
                self.format_and_write(pdf, prep.get("talking_points", []), title="Talking Points")
                self.format_and_write(pdf, prep.get("competitor_mentions", []), title="Competitor Mentions")

                self.format_and_write(pdf, data.get("follow_up_draft", ""), title="Follow-Up Draft")

            # --- Existing Coaching Report (skill_scores) ---
            elif "skill_scores" in data:
                pdf.set_font("DejaVu", "B", 14)
                pdf.cell(0, 10, "Skill Scores", ln=True)
                self.render_table(pdf, data["skill_scores"])
                # radar chart can be embedded here

                for k, v in data.items():
                    if k == "skill_scores": 
                        continue
                    self.format_and_write(pdf, v, title=str(k))

            # --- Deal Execution Summary ---
            elif "overview" in data:
                pdf.set_font("DejaVu", "B", 16)
                pdf.cell(0, 10, "Deal Execution Summary", ln=True)

                self.format_and_write(pdf, data.get("overview", ""), title="Overview")
                self.format_and_write(pdf, data.get("outline", ""), title="Outline")
                self.format_and_write(pdf, data.get("notes", []), title="Notes")
                self.format_and_write(pdf, data.get("action_items", {}), title="Action Items")

            else:
                # fallback
                self.format_and_write(pdf, data)







        else:
            self.format_and_write(pdf, data)

        pdf.output(path)
        return path

