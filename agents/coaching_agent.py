import os
import json
from typing import Dict, Any
from langchain.chat_models import ChatOpenAI
from langchain.chains import LLMChain
from prompts.coaching_agent_prompt import coaching_agent_prompt
from utils.transcript_loader import TranscriptLoader
from utils.transcript_chunker import TranscriptChunker
from utils.excel_handler import ExcelHandler
from utils.output_writer import OutputWriter


class CoachingAgent:
    def __init__(self, 
                 crm_path: str = "crm_data.xlsx", 
                 meetings_path: str = "meetings.xlsx",
                 model: str = "gpt-3.5-turbo"):
        """Initialize Coaching Agent with helpers and OpenAI model."""
        self.transcript_loader = TranscriptLoader(meetings_path)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.writer = OutputWriter()
        self.chunker = TranscriptChunker(model_name=model)
        self.llm = ChatOpenAI(
            model=model,
            temperature=0.3,
            openai_api_key=os.getenv("OPENAI_API_KEY")
        )
        self.chain = LLMChain(prompt=coaching_agent_prompt, llm=self.llm)

    def get_roles_from_meeting(self, meeting_id: str, client_name: str):
        """Parse rep vs client roles dynamically from meetings.xlsx Participants column."""
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
        participants = str(meeting_info.get("Participants", ""))

        reps, clients = [], []
        for p in participants.split(","):
            p = p.strip()
            if "(" in p and ")" in p:
                name, role = p.rsplit("(", 1)
                role = role.replace(")", "").strip()
                name = name.strip()
                if role.lower() in ["dice", "rep", "sales", "seller"]:
                    reps.append(name)
                else:
                    clients.append(name)
            else:
                clients.append(p)
        return reps, clients

    def process(self, meeting_id: str, client_name: str) -> Dict[str, Any]:
        """
        Analyze past meeting transcript and generate coaching report.
        """
        # --- Step 1: Load transcript ---
        dialogue, _ = self.transcript_loader.load_transcript(meeting_id, client_name)

        # --- Step 2: Detect roles dynamically ---
        reps, clients = self.get_roles_from_meeting(meeting_id, client_name)
        print(f"[INFO] Detected reps: {reps}, clients: {clients}")

        # --- Step 3: Chunk transcript ---
        chunks = self.chunker.chunk_transcript(dialogue)

        all_scores, reports = [], []
        for chunk in chunks:
            response = self.chain.invoke({
                "transcript": chunk,
                "reps": ", ".join(reps),
                "clients": ", ".join(clients),
                
            })
            text = response["text"] if isinstance(response, dict) else response
            reports.append(text)

            try:
                parsed = json.loads(text)
                if "skill_scores" in parsed:
                    all_scores.append(parsed["skill_scores"])
            except Exception:
                pass

        # --- Step 4: Merge scores ---
        final_scores = {}
        if all_scores:
            for key in all_scores[0].keys():
                values = [s[key] for s in all_scores if key in s]
                final_scores[key] = int(sum(values) / len(values)) if values else 0

        report = {
            "Meeting_ID": meeting_id,
            "Client": client_name,
            "Reports": reports,
            "Averaged_Scores": final_scores
        }

        # --- Step 5: Save outputs ---
        self.writer.append_to_excel(report, "coaching_reports.xlsx", sheet_name="Coaching")

        if final_scores:
            self.writer.save_radar_chart(
                final_scores,
                f"Skill Report - {client_name} {meeting_id}",
                f"{client_name}_{meeting_id}_skills.png"
            )

        return report
