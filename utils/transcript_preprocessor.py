import json
from typing import List, Dict, Tuple
from utils.transcript_loader import TranscriptLoader
from utils.excel_handler import ExcelHandler
from utils.transcript_chunker import TranscriptChunker


class TranscriptPreprocessor:
    """
    Preprocess transcripts once and return structured dialogue + roles.
    """

    def __init__(self, meetings_path: str = "meetings.xlsx", crm_path: str = "crm_data.xlsx"):
        self.loader = TranscriptLoader(meetings_path)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.chunker = TranscriptChunker()

    def preprocess(self, meeting_id: str, client_name: str) -> Tuple[List[Dict[str, str]], List[str], List[str]]:
        """
        Loads transcript, structures it, and extracts roles automatically.
        Returns: (dialogue, reps, clients)
        """
        # --- Step 1: Load transcript
        dialogue, _ = self.loader.load_transcript(meeting_id, client_name)

        # --- Step 2: Get participants from Excel
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
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

        return dialogue, reps, clients
