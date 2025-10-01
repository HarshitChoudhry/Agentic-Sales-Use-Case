import os
import json
import pandas as pd
from typing import List, Dict, Any, Tuple, Optional

class TranscriptLoader:
    """
    Utility class to load, parse, and format meeting transcripts stored as JSON.
    """

    def __init__(self, meetings_excel: str = "meetings.xlsx"):
        """
        Initialize with path to meetings.xlsx so we can map Meeting_ID -> transcript path.
        """
        self.meetings_excel = meetings_excel

    def get_transcript_path(self, meeting_id: str, client_name: str) -> Optional[str]:
        """
        Get transcript file path for a given meeting_id from meetings.xlsx (client sheet).
        """
        try:
            # Load the specific client sheet
            df = pd.read_excel(self.meetings_excel, sheet_name=client_name)
            row = df[df["Meeting_ID"] == meeting_id]

            if row.empty:
                return None

            return str(row.iloc[0]["Transcript_File"]) if pd.notna(row.iloc[0]["Transcript_File"]) else None
        except Exception as e:
            print(f"[ERROR] Could not read meetings.xlsx for {client_name}, meeting {meeting_id}: {e}")
            return None

    def load_json(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Load transcript JSON file and return raw list of utterances.
        Each utterance has: sentence, startTime, endTime, speaker_name, speaker_id
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Transcript file not found: {file_path}")

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format in {file_path}: {e}")

    def to_structured_dialogue(self, data: List[Dict[str, Any]]) -> List[Dict[str, str]]:
        """
        Convert raw transcript JSON into structured dialogue.
        """
        dialogue = []
        for entry in data:
            dialogue.append({
                "speaker": entry.get("speaker_name", "Unknown"),
                "start": entry.get("startTime", ""),
                "end": entry.get("endTime", ""),
                "sentence": entry.get("sentence", "").strip()
            })
        return dialogue

    def to_readable_text(self, dialogue: List[Dict[str, str]]) -> str:
        """
        Convert structured dialogue into a human-readable transcript string.
        """
        lines = []
        for entry in dialogue:
            speaker = entry["speaker"]
            start = entry["start"]
            sentence = entry["sentence"]
            if sentence:
                lines.append(f"{speaker} - {start}\n{sentence}\n")
        return "\n".join(lines)

    def load_transcript(
        self, meeting_id: str, client_name: str
    ) -> Tuple[List[Dict[str, str]], str]:
        """
        Load transcript by meeting_id + client_name.
        Returns (structured_dialogue, readable_text).
        """
        file_path = self.get_transcript_path(meeting_id, client_name)
        if not file_path:
            raise FileNotFoundError(f"No transcript path found for {meeting_id} in {client_name} sheet.")

        data = self.load_json(file_path)
        dialogue = self.to_structured_dialogue(data)
        readable = self.to_readable_text(dialogue)
        return dialogue, readable
