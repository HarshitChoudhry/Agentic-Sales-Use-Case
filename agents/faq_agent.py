import os
import json
from typing import Dict, Any, List
from langchain.chat_models import ChatOpenAI
from langchain.chains import LLMChain
from prompts.faq_agent_prompt import faq_agent_prompt
from utils.transcript_loader import TranscriptLoader
from utils.transcript_chunker import TranscriptChunker
from utils.excel_handler import ExcelHandler
from utils.output_writer import OutputWriter


class FAQAgent:
    def __init__(self, 
                 crm_path: str = "crm_data.xlsx", 
                 meetings_path: str = "meetings.xlsx",
                 model: str = "gpt-3.5-turbo"):
        """Initialize FAQ Agent with helpers and OpenAI model."""
        self.transcript_loader = TranscriptLoader(meetings_path)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.writer = OutputWriter()
        self.chunker = TranscriptChunker(model_name=model)
        self.llm = ChatOpenAI(
            model=model,
            temperature=0.3,
            openai_api_key=os.getenv("OPENAI_API_KEY")
        )
        self.chain = LLMChain(prompt=faq_agent_prompt, llm=self.llm)

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

    def extract_client_lines(self, transcript: List[Dict[str, str]], client_speakers: List[str]) -> List[str]:
        """
        Extract ALL client-side utterances (not just questions).
        Let the LLM decide which ones are FAQ-worthy.
        """
        utterances = []
        norm_clients = [c.lower() for c in client_speakers]

        for entry in transcript:
            text = entry.get("sentence", "").strip()
            speaker = entry.get("speaker", "").lower()
            if not text:
                continue

            if any(c in speaker for c in norm_clients):
                utterances.append(f"{speaker}: {text}")

        return utterances

    def process(self, meeting_id: str, client_name: str) -> List[Dict[str, Any]]:
        """
        Process FAQ extraction for a past meeting.
        """
        # --- Step 1: Load transcript ---
        dialogue, _ = self.transcript_loader.load_transcript(meeting_id, client_name)

        # --- Step 2: Detect roles dynamically ---
        reps, clients = self.get_roles_from_meeting(meeting_id, client_name)
        print(f"[INFO] Detected reps: {reps}, clients: {clients}")

        # --- Step 3: Chunk if transcript is too long ---
        if self.chunker.count_tokens(str(dialogue)) > 12000:
            chunks = self.chunker.chunk_transcript(dialogue)
        else:
            chunks = [dialogue]

        results = []
        for chunk in chunks:
            # Extract ALL client utterances from this chunk
            utterances = self.extract_client_lines(chunk if isinstance(chunk, list) else [], clients)

            for line in utterances:
                response = self.chain.invoke({
                    "question": line,
                    "transcript": str(chunk),
                    "reps": ", ".join(reps),
                    "clients": ", ".join(clients)
                })
                text = response["text"] if isinstance(response, dict) else response

                faq_entry = {
                    "Meeting_ID": meeting_id,
                    "Client": client_name,
                    "Client_Utterance": line,
                    "Analysis": text
                }
                results.append(faq_entry)

                # Save incrementally
                self.writer.append_to_excel(faq_entry, "faq_knowledge.xlsx", sheet_name="FAQ")

        return results
