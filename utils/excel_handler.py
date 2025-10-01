import os
import pandas as pd
from typing import Dict, Any, List, Optional

class ExcelHandler:
    """
    Utility class to manage crm_data.xlsx and meetings.xlsx (multi-sheet).
    """

    def __init__(self, crm_path: str = "crm_data.xlsx", meetings_path: str = "meetings.xlsx"):
        self.crm_path = crm_path
        self.meetings_path = meetings_path

    # ----------------- CRM Methods -----------------
    def get_deal_info(self, deal_id: str) -> Optional[Dict[str, Any]]:
        """Fetch CRM row by Deal_ID."""
        try:
            df = pd.read_excel(self.crm_path, sheet_name="CRM_Data")
            row = df[df["Deal_ID"] == deal_id]
            if row.empty:
                return None
            return row.iloc[0].to_dict()
        except Exception as e:
            print(f"[ERROR] Could not read CRM data: {e}")
            return None

    def update_deal_info(self, deal_id: str, updates: Dict[str, Any]) -> bool:
        """Update specific fields in CRM by Deal_ID."""
        try:
            df = pd.read_excel(self.crm_path, sheet_name="CRM_Data")
            idx = df.index[df["Deal_ID"] == deal_id].tolist()
            if not idx:
                return False
            for col, val in updates.items():
                if col in df.columns:
                    df.loc[idx[0], col] = val
            df.to_excel(self.crm_path, sheet_name="CRM_Data", index=False)
            return True
        except Exception as e:
            print(f"[ERROR] Could not update CRM data: {e}")
            return False

    # ----------------- Meeting Methods -----------------
    def get_meeting_info(self, meeting_id: str, client_name: str) -> Optional[Dict[str, Any]]:
        """Fetch meeting row by Meeting_ID and client sheet."""
        try:
            df = pd.read_excel(self.meetings_path, sheet_name=client_name)
            row = df[df["Meeting_ID"] == meeting_id]
            if row.empty:
                return None
            return row.iloc[0].to_dict()
        except Exception as e:
            print(f"[ERROR] Could not read meeting data: {e}")
            return None
    def get_deal_by_client(self, client_name: str):
        """Return deal info by client name (case-insensitive)."""
        import pandas as pd
        df = pd.read_excel(self.crm_path)
        match = df[df["Client"].str.lower() == client_name.lower()]
        if not match.empty:
            return match.iloc[0].to_dict()
        return None

    def list_meetings(self, client_name: str, status: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        List all meetings for a client.
        Optionally filter by status = 'past' or 'upcoming'.
        """
        try:
            df = pd.read_excel(self.meetings_path, sheet_name=client_name)
            if status:
                df = df[df["Status"].str.lower() == status.lower()]
            return df.to_dict(orient="records")
        except Exception as e:
            print(f"[ERROR] Could not list meetings: {e}")
            return []

    def update_meeting_info(self, meeting_id: str, client_name: str, updates: Dict[str, Any]) -> bool:
        """Update specific fields in a meeting row."""
        try:
            # Load all sheets
            xls = pd.ExcelFile(self.meetings_path)
            writer = pd.ExcelWriter(self.meetings_path, engine="xlsxwriter")

            for sheet in xls.sheet_names:
                df = pd.read_excel(self.meetings_path, sheet_name=sheet)
                if sheet == client_name:
                    idx = df.index[df["Meeting_ID"] == meeting_id].tolist()
                    if idx:
                        for col, val in updates.items():
                            if col in df.columns:
                                df.loc[idx[0], col] = val
                df.to_excel(writer, sheet_name=sheet, index=False)

            writer.close()
            return True
        except Exception as e:
            print(f"[ERROR] Could not update meeting info: {e}")
            return False

