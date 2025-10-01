import os
import json
from typing import Dict, Any
from langchain.chat_models import ChatOpenAI
from langchain.chains import LLMChain
from prompts.meeting_prep_prompt import meeting_prep_prompt
from utils.transcript_loader import TranscriptLoader
from utils.excel_handler import ExcelHandler
from utils.output_writer import OutputWriter
from utils.transcript_chunker import TranscriptChunker


class MeetingPreparationAgent:
    def __init__(self,
                 crm_path: str = "crm_data.xlsx",
                 meetings_path: str = "meetings.xlsx",
                 model: str = "gpt-3.5-turbo"):
        self.transcript_loader = TranscriptLoader(meetings_path)
        self.chunker = TranscriptChunker(model_name=model)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.writer = OutputWriter()
        self.llm = ChatOpenAI(
            model=model,
            temperature=0.3,
            openai_api_key=os.getenv("OPENAI_API_KEY"),
        )
        self.chain = LLMChain(prompt=meeting_prep_prompt, llm=self.llm)

    def get_roles_for_meeting(self, meeting_id: str, client_name: str):
        """Extract reps and clients from Participants column."""
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
        if not meeting_info:
            return [], []
        participants = str(meeting_info.get("Participants", ""))
        reps, clients = [], []
        for p in participants.split(","):
            p = p.strip()
            if "(" in p and ")" in p:
                name, role = p.rsplit("(", 1)
                role = role.replace(")", "").strip().lower()
                name = name.strip()
                if role in ("dice", "rep", "sales", "seller"):
                    reps.append(name)
                else:
                    clients.append(name)
            else:
                clients.append(p)
        return reps, clients

    def prepare_meeting(self, meeting_id: str, client_name: str) -> Dict[str, Any]:
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
        if not meeting_info:
            raise ValueError(f"No meeting found for {meeting_id} in {client_name}")

        # CRM lookup
        deal_id = meeting_info.get("Deal_ID")
        if not deal_id or str(deal_id).lower() == "nan":
            crm_info = self.excel_handler.get_deal_by_client(client_name)
            if crm_info:
                deal_id = crm_info["Deal_ID"]
                print(f"[INFO] Auto-assigned Deal_ID {deal_id} for {client_name}")
            else:
                raise ValueError(f"No CRM info for client {client_name}")
        crm_info = self.excel_handler.get_deal_info(deal_id)

        # Past meeting summaries
        past_meetings = self.excel_handler.list_meetings(client_name, status="past")
        past_context = "\n".join(
            [f"{m['Title']} ({m['Date']}): {m.get('Notes','')}" for m in past_meetings]
        ) if past_meetings else "No past context available."

        # Roles
        reps, clients = self.get_roles_for_meeting(meeting_id, client_name)

        # LLM call
        response = self.chain.invoke({
            "deal_info": json.dumps(crm_info, default=str),
            "past_meetings": past_context,
            "reps": ", ".join(reps) if reps else "None",
            "clients": ", ".join(clients) if clients else "None",
            "upcoming_meeting": json.dumps(meeting_info, default=str),
        })
        response_text = response["text"] if isinstance(response, dict) else str(response)

        result = {"Meeting_ID": meeting_id, "Client": client_name, "Output": response_text}

        # Save
        self.writer.append_to_excel(result, "meeting_guidance.xlsx", sheet_name="Guidance")
        self.writer.write_to_pdf(
            response_text,
            f"{client_name}_{meeting_id}_prep.pdf",
            subfolder="prep_packs",
            meta={
                "client": client_name,
                "meeting_id": meeting_id,
                "date": meeting_info.get("Date", ""),
                "report_type": "Meeting Preparation Pack",
            },
        )

        return result
