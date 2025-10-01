import os
import json
from typing import Dict, Any, List
from langchain.chat_models import ChatOpenAI
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
from prompts.deal_exec_prompt import deal_exec_prompt
from utils.transcript_loader import TranscriptLoader
from utils.transcript_chunker import TranscriptChunker
from utils.excel_handler import ExcelHandler
from utils.output_writer import OutputWriter


class DealExecutionAgent:
    def __init__(self,
                 crm_path: str = "crm_data.xlsx",
                 meetings_path: str = "meetings.xlsx",
                 model: str = "gpt-3.5-turbo"):
        """Initialize Deal Execution Agent with helpers and OpenAI model."""
        self.transcript_loader = TranscriptLoader(meetings_path)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.writer = OutputWriter()
        self.chunker = TranscriptChunker(model_name=model)
        self.llm = ChatOpenAI(
            model=model,
            temperature=0.3,
            openai_api_key=os.getenv("OPENAI_API_KEY")
        )
        self.chain = LLMChain(prompt=deal_exec_prompt, llm=self.llm)

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
        Generate post-meeting Deal Execution Summary for a past meeting.
        """
        # --- Step 1: Load transcript ---
        dialogue, _ = self.transcript_loader.load_transcript(meeting_id, client_name)

        # --- Step 2: Detect roles dynamically ---
        reps, clients = self.get_roles_from_meeting(meeting_id, client_name)
        print(f"[INFO] Detected reps: {reps}, clients: {clients}")

        # --- Step 3: Chunk transcript ---
        chunks = self.chunker.chunk_transcript(dialogue)

        # --- Step 4: Summarize with reps/clients context ---
        final_summary = self.chunker.summarize_chunks(
            chunks,
            self.chain,
            extra_inputs={
                "reps": ", ".join(reps) if reps else "None",
                "clients": ", ".join(clients) if clients else "None"
            }
        )

        # --- Step 5: Parse JSON ---
        try:
            parsed = json.loads(final_summary)
        except Exception:
            parsed = {"raw_text": final_summary}

        result = {"Meeting_ID": meeting_id, "Client": client_name, "Output": parsed}

        # --- Step 6: Save Excel Log ---
        self.writer.append_to_excel(result, "deal_execution.xlsx", sheet_name="Execution")

        # --- Step 7: Save PDF Report with header ---
        self.writer.write_to_pdf(
            parsed,
            f"{client_name}_{meeting_id}_summary.pdf",
            subfolder="deal_summaries",
            meta={
                "client": client_name,
                "meeting_id": meeting_id,
                "date": self.excel_handler.get_meeting_info(meeting_id, client_name).get("Date", ""),
                "report_type": "Post-Meeting Execution Summary"
            }
        )

        return result

    def process_account(self, client_name: str) -> Dict[str, Any]:
        """
        Combine multiple meeting-level summaries into one account-level summary.
        """
        past_meetings = self.excel_handler.list_meetings(client_name, status="past")
        meeting_summaries: List[Dict[str, Any]] = []

        for m in past_meetings:
            meeting_id = m["Meeting_ID"]
            try:
                result = self.process(meeting_id, client_name)
                meeting_summaries.append({
                    "Meeting_ID": meeting_id,
                    "Summary": result["Output"]
                })
            except Exception as e:
                print(f"⚠️ Deal Execution failed for {client_name} {meeting_id}: {e}")

        if not meeting_summaries:
            return {"Client": client_name, "Account_Summary": "No past meetings found."}

        # --- Combine summaries into account-level ---
        combine_prompt = PromptTemplate(
            input_variables=["meeting_summaries"],
            template="""
You are the Deal Execution Agent.
You will combine multiple meeting-level summaries into ONE coherent account-level narrative.

Meeting Summaries:
{meeting_summaries}

Return in JSON:
{{
  "overall_overview": "Narrative covering everything that has happened so far with the account",
  "timeline": ["Key milestones across meetings"],
  "consolidated_notes": ["Merged notes across meetings"],
  "open_action_items": {{
    "sales_rep": ["..."],
    "client": ["..."]
  }}
}}
"""
        )

        combine_chain = LLMChain(prompt=combine_prompt, llm=self.llm)
        combined_summary = combine_chain.invoke(
            {"meeting_summaries": json.dumps(meeting_summaries, indent=2)}
        )
        text = combined_summary["text"] if isinstance(combined_summary, dict) else combined_summary

        try:
            parsed = json.loads(text)
        except Exception:
            parsed = {"raw_text": text}

        result = {"Client": client_name, "Account_Summary": parsed}

        # --- Save Excel Log ---
        self.writer.append_to_excel(result, "deal_execution.xlsx", sheet_name="Account_Level")

        # --- Save Account-level PDF ---
        self.writer.write_to_pdf(
            parsed,
            f"{client_name}_account_summary.pdf",
            subfolder="deal_summaries",
            meta={
                "client": client_name,
                "meeting_id": "ALL",
                "date": "",
                "report_type": "Account-Level Execution Summary"
            }
        )

        return result
