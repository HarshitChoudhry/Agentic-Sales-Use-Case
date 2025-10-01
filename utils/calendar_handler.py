# # # # # import os
# # # # # import datetime
# # # # # import pandas as pd
# # # # # from typing import List, Dict, Any
# # # # # from difflib import SequenceMatcher
# # # # # from google.oauth2.credentials import Credentials
# # # # # from googleapiclient.discovery import build
# # # # # from dotenv import load_dotenv

# # # # # # Load environment variables
# # # # # load_dotenv()


# # # # # def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.6) -> str:
# # # # #     """Find the best fuzzy match for target among candidates."""
# # # # #     best_match, score = None, 0
# # # # #     for c in candidates:
# # # # #         ratio = SequenceMatcher(None, target.lower(), c.lower()).ratio()
# # # # #         if ratio > score:
# # # # #             best_match, score = c, ratio
# # # # #     return best_match if score >= threshold else None


# # # # # class CalendarHandler:
# # # # #     """
# # # # #     Utility to fetch upcoming meetings from Google Calendar 
# # # # #     and sync them into meetings.xlsx client sheets.
# # # # #     """

# # # # #     def __init__(self, meetings_file: str = "meetings.xlsx"):
# # # # #         self.meetings_file = meetings_file
# # # # #         self.service = self.authenticate()

# # # # #     def authenticate(self):
# # # # #         """Authenticate with Google Calendar using .env credentials."""
# # # # #         creds = Credentials(
# # # # #             token=None,
# # # # #             refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
# # # # #             token_uri="https://oauth2.googleapis.com/token",
# # # # #             client_id=os.getenv("GOOGLE_CLIENT_ID"),
# # # # #             client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
# # # # #             scopes=["https://www.googleapis.com/auth/calendar.readonly"],
# # # # #         )
# # # # #         return build("calendar", "v3", credentials=creds)

# # # # #     def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
# # # # #         """Fetch upcoming events within the next N days."""
# # # # #         now = datetime.datetime.utcnow().isoformat() + "Z"
# # # # #         future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

# # # # #         events_result = (
# # # # #             self.service.events()
# # # # #             .list(
# # # # #                 calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
# # # # #                 timeMin=now,
# # # # #                 timeMax=future,
# # # # #                 singleEvents=True,
# # # # #                 orderBy="startTime",
# # # # #             )
# # # # #             .execute()
# # # # #         )

# # # # #         events = events_result.get("items", [])
# # # # #         meetings = []

# # # # #         for event in events:
# # # # #             title = event.get("summary", "Untitled Meeting")
# # # # #             start = event["start"].get("dateTime", event["start"].get("date"))

# # # # #             attendees = []
# # # # #             if "attendees" in event:
# # # # #                 for att in event["attendees"]:
# # # # #                     name = att.get("displayName") or att.get("email", "Unknown")
# # # # #                     attendees.append(name)

# # # # #             meetings.append({
# # # # #                 "title": title,
# # # # #                 "date": start,
# # # # #                 "participants": attendees,
# # # # #                 "event_id": event.get("id", "")
# # # # #             })

# # # # #         return meetings

# # # # #     def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
# # # # #         """
# # # # #         Add fetched events into meetings.xlsx.
# # # # #         Uses fuzzy matching for client sheets and splits participants into reps vs clients.
# # # # #         Returns: list of updated client sheet names.
# # # # #         """
# # # # #         updated_clients = []
# # # # #         try:
# # # # #             xls = pd.ExcelFile(self.meetings_file)
# # # # #             client_sheets = xls.sheet_names

# # # # #             writer = pd.ExcelWriter(self.meetings_file, engine="openpyxl", mode="a", if_sheet_exists="overlay")

# # # # #             for event in events:
# # # # #                 title = event["title"]
# # # # #                 client_name = fuzzy_match(title, client_sheets)

# # # # #                 if not client_name:
# # # # #                     print(f"[WARN] No client sheet matched for event '{title}'")
# # # # #                     continue

# # # # #                 df = pd.read_excel(self.meetings_file, sheet_name=client_name)

# # # # #                 # Avoid duplicates
# # # # #                 if "Meeting_ID" in df.columns and event["event_id"] in df["Meeting_ID"].values:
# # # # #                     continue

# # # # #                 # --- Participant Role Detection (patched for your use-case) ---
# # # # #                 reps, clients = [], []
# # # # #                 for p in event["participants"]:
# # # # #                     if "aakashgupta" in p.lower():  # treat as rep
# # # # #                         reps.append(p)
# # # # #                     else:
# # # # #                         clients.append(p)

# # # # #                 new_row = {
# # # # #                     "Meeting_ID": event["event_id"],
# # # # #                     "Deal_ID": "",  # can be filled later
# # # # #                     "Title": event["title"],
# # # # #                     "Date": event["date"],
# # # # #                     "Participants": f"Sales Reps: {', '.join(reps)} | Clients: {', '.join(clients)}",
# # # # #                     "Status": "upcoming",
# # # # #                     "Transcript_File": "",
# # # # #                     "Notes": "Imported from Google Calendar"
# # # # #                 }

# # # # #                 df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
# # # # #                 df.to_excel(writer, sheet_name=client_name, index=False)

# # # # #                 updated_clients.append(client_name)
# # # # #                 print(f"✅ Added new meeting for {client_name}: {title}")

# # # # #             writer.close()
# # # # #             return list(set(updated_clients))

# # # # #         except Exception as e:
# # # # #             print(f"[ERROR] Failed to update meetings.xlsx: {e}")
# # # # #             return []




# # # # import os
# # # # import datetime
# # # # import pandas as pd
# # # # from typing import List, Dict, Any
# # # # from difflib import SequenceMatcher
# # # # from google.oauth2.credentials import Credentials
# # # # from googleapiclient.discovery import build
# # # # from dotenv import load_dotenv

# # # # # Load environment variables
# # # # load_dotenv()


# # # # def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.3) -> str:
# # # #     """Find the best fuzzy match for target among candidates."""
# # # #     best_match, score = None, 0
# # # #     for c in candidates:
# # # #         ratio = SequenceMatcher(None, target.lower(), c.lower()).ratio()
# # # #         if ratio > score:
# # # #             best_match, score = c, ratio
# # # #     return best_match if score >= threshold else None


# # # # class CalendarHandler:
# # # #     """
# # # #     Utility to fetch upcoming meetings from Google Calendar 
# # # #     and sync them into meetings.xlsx client sheets.
# # # #     """

# # # #     def __init__(self, meetings_file: str = "meetings.xlsx"):
# # # #         self.meetings_file = meetings_file
# # # #         self.service = self.authenticate()

# # # #     def authenticate(self):
# # # #         """Authenticate with Google Calendar using .env credentials."""
# # # #         creds = Credentials(
# # # #             token=None,
# # # #             refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
# # # #             token_uri="https://oauth2.googleapis.com/token",
# # # #             client_id=os.getenv("GOOGLE_CLIENT_ID"),
# # # #             client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
# # # #             scopes=["https://www.googleapis.com/auth/calendar.readonly"],
# # # #         )
# # # #         return build("calendar", "v3", credentials=creds)

# # # #     def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
# # # #         """Fetch upcoming events within the next N days."""
# # # #         now = datetime.datetime.utcnow().isoformat() + "Z"
# # # #         future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

# # # #         events_result = (
# # # #             self.service.events()
# # # #             .list(
# # # #                 calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
# # # #                 timeMin=now,
# # # #                 timeMax=future,
# # # #                 singleEvents=True,
# # # #                 orderBy="startTime",
# # # #             )
# # # #             .execute()
# # # #         )

# # # #         events = events_result.get("items", [])
# # # #         meetings = []

# # # #         for event in events:
# # # #             title = event.get("summary", "Untitled Meeting")
# # # #             start = event["start"].get("dateTime", event["start"].get("date"))

# # # #             attendees = []
# # # #             if "attendees" in event:
# # # #                 for att in event["attendees"]:
# # # #                     name = att.get("displayName") or att.get("email", "Unknown")
# # # #                     attendees.append(name)

# # # #             meetings.append({
# # # #                 "title": title,
# # # #                 "date": start,
# # # #                 "participants": attendees,
# # # #                 "event_id": event.get("id", "")
# # # #             })

# # # #         return meetings

# # # #     def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
# # # #         """
# # # #         Add fetched events into meetings.xlsx.
# # # #         Uses fuzzy matching for client sheets and splits participants into reps vs clients.
# # # #         Returns: list of updated client sheet names.
# # # #         """
# # # #         updated_clients = []
# # # #         try:
# # # #             xls = pd.ExcelFile(self.meetings_file)
# # # #             client_sheets = xls.sheet_names
# # # #             all_sheets = {sheet: pd.read_excel(self.meetings_file, sheet_name=sheet) for sheet in client_sheets}

# # # #             for event in events:
# # # #                 title = event["title"]
# # # #                 client_name = fuzzy_match(title, client_sheets)

# # # #                 if not client_name:
# # # #                     print(f"[WARN] No client sheet matched for event '{title}'")
# # # #                     continue

# # # #                 df = all_sheets[client_name]

# # # #                 # Avoid duplicates
# # # #                 if "Meeting_ID" in df.columns and event["event_id"] in df["Meeting_ID"].values:
# # # #                     continue

# # # #                 # --- Participant Role Detection (patched for your case) ---
# # # #                 reps, clients = [], []
# # # #                 for p in event["participants"]:
# # # #                     if "aakashgupta" in p.lower():
# # # #                         reps.append(p)
# # # #                     else:
# # # #                         clients.append(p)

# # # #                 new_row = {
# # # #                     "Meeting_ID": event["event_id"],
# # # #                     "Deal_ID": "",  # can be filled later
# # # #                     "Title": event["title"],
# # # #                     "Date": event["date"],
# # # #                     "Participants": f"Sales Reps: {', '.join(reps)} | Clients: {', '.join(clients)}",
# # # #                     "Status": "upcoming",
# # # #                     "Transcript_File": "",
# # # #                     "Notes": "Imported from Google Calendar"
# # # #                 }

# # # #                 df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
# # # #                 all_sheets[client_name] = df
# # # #                 updated_clients.append(client_name)
# # # #                 print(f"✅ Added new meeting for {client_name}: {title}")

# # # #             # Write all sheets back at once (avoids file-lock issues)
# # # #             with pd.ExcelWriter(self.meetings_file, engine="openpyxl", mode="w") as writer:
# # # #                 for sheet, df in all_sheets.items():
# # # #                     df.to_excel(writer, sheet_name=sheet, index=False)

# # # #             return list(set(updated_clients))

# # # #         except Exception as e:
# # # #             print(f"[ERROR] Failed to update meetings.xlsx: {e}")
# # # #             return []









# # # import os
# # # import datetime
# # # import pandas as pd
# # # from typing import List, Dict, Any
# # # from difflib import SequenceMatcher
# # # from google.oauth2.credentials import Credentials
# # # from googleapiclient.discovery import build
# # # from dotenv import load_dotenv

# # # # Load environment variables
# # # load_dotenv()


# # # def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.6) -> str:
# # #     """Find the best fuzzy match for target among candidates."""
# # #     best_match, score = None, 0
# # #     for c in candidates:
# # #         ratio = SequenceMatcher(None, target.lower(), c.lower()).ratio()
# # #         if ratio > score:
# # #             best_match, score = c, ratio
# # #     return best_match if score >= threshold else None


# # # class CalendarHandler:
# # #     """
# # #     Utility to fetch upcoming meetings from Google Calendar 
# # #     and sync them into meetings.xlsx client sheets.
# # #     """

# # #     def __init__(self, meetings_file: str = "meetings.xlsx", crm_file: str = "crm_data.xlsx"):
# # #         self.meetings_file = meetings_file
# # #         self.crm_file = crm_file
# # #         self.service = self.authenticate()

# # #     def authenticate(self):
# # #         """Authenticate with Google Calendar using .env credentials."""
# # #         creds = Credentials(
# # #             token=None,
# # #             refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
# # #             token_uri="https://oauth2.googleapis.com/token",
# # #             client_id=os.getenv("GOOGLE_CLIENT_ID"),
# # #             client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
# # #             scopes=["https://www.googleapis.com/auth/calendar.readonly"],
# # #         )
# # #         return build("calendar", "v3", credentials=creds)

# # #     def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
# # #         """Fetch upcoming events within the next N days."""
# # #         now = datetime.datetime.utcnow().isoformat() + "Z"
# # #         future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

# # #         events_result = (
# # #             self.service.events()
# # #             .list(
# # #                 calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
# # #                 timeMin=now,
# # #                 timeMax=future,
# # #                 singleEvents=True,
# # #                 orderBy="startTime",
# # #             )
# # #             .execute()
# # #         )

# # #         events = events_result.get("items", [])
# # #         meetings = []

# # #         for event in events:
# # #             title = event.get("summary", "Untitled Meeting")
# # #             start = event["start"].get("dateTime", event["start"].get("date"))

# # #             attendees = []
# # #             if "attendees" in event:
# # #                 for att in event["attendees"]:
# # #                     name = att.get("displayName") or att.get("email", "Unknown")
# # #                     attendees.append(name)

# # #             meetings.append({
# # #                 "title": title,
# # #                 "date": start,
# # #                 "participants": attendees
# # #             })

# # #         return meetings

# # #     def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
# # #         """
# # #         Add fetched events into meetings.xlsx.
# # #         - Auto-increments Meeting_IDs per client.
# # #         - Looks up Deal_ID from crm_data.xlsx.
# # #         Returns: list of updated client sheet names.
# # #         """
# # #         updated_clients = []
# # #         try:
# # #             # Load CRM
# # #             crm_df = pd.read_excel(self.crm_file)

# # #             # Load meetings.xlsx sheets
# # #             xls = pd.ExcelFile(self.meetings_file)
# # #             client_sheets = xls.sheet_names
# # #             all_sheets = {sheet: pd.read_excel(self.meetings_file, sheet_name=sheet) for sheet in client_sheets}

# # #             for event in events:
# # #                 title = event["title"]

# # #                 # Normalize + fuzzy match
# # #                 def normalize(n): return n.lower().replace("_", " ").replace("-", " ").replace("technologies", "").strip()
# # #                 client_name = fuzzy_match(normalize(title), [normalize(s) for s in client_sheets])

# # #                 if not client_name:
# # #                     print(f"[WARN] No client sheet matched for event '{title}'")
# # #                     continue

# # #                 # Get actual sheet name from normalized mapping
# # #                 for sheet in client_sheets:
# # #                     if normalize(sheet) == client_name:
# # #                         client_name = sheet
# # #                         break

# # #                 df = all_sheets[client_name]

# # #                 # --- Auto-increment Meeting_ID ---
# # #                 if "Meeting_ID" in df.columns and not df.empty:
# # #                     last_id = df["Meeting_ID"].dropna().iloc[-1]
# # #                     try:
# # #                         next_id_num = int(str(last_id).replace("M", "")) + 1
# # #                     except ValueError:
# # #                         next_id_num = len(df) + 1
# # #                     meeting_id = f"M{next_id_num:03d}"
# # #                 else:
# # #                     meeting_id = "M001"

# # #                 # --- Fetch Deal_ID from CRM ---
# # #                 crm_df["norm"] = crm_df["Client"].str.lower().str.replace("-", " ").str.replace("_", " ")
# # #                 crm_match = crm_df[crm_df["norm"] == normalize(client_name)]
# # #                 deal_id = crm_match["Deal_ID"].iloc[0] if not crm_match.empty else ""

# # #                 # --- Role Detection (simple rule) ---
# # #                 reps, clients = [], []
# # #                 for p in event["participants"]:
# # #                     if "aakashgupta" in p.lower():  # your rep name
# # #                         reps.append(p)
# # #                     else:
# # #                         clients.append(p)

# # #                 new_row = {
# # #                     "Meeting_ID": meeting_id,
# # #                     "Deal_ID": deal_id,
# # #                     "Title": event["title"],
# # #                     "Date": event["date"],
# # #                     "Participants": f"Sales Reps: {', '.join(reps)} | Clients: {', '.join(clients)}",
# # #                     "Status": "upcoming",
# # #                     "Transcript_File": "",
# # #                     "Notes": "Imported from Google Calendar"
# # #                 }

# # #                 df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
# # #                 all_sheets[client_name] = df
# # #                 updated_clients.append(client_name)
# # #                 print(f"✅ Added new meeting for {client_name}: {title} -> {meeting_id}")

# # #             # --- Write back ONLY updated sheets ---
# # #             with pd.ExcelWriter(self.meetings_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
# # #                 for client in set(updated_clients):
# # #                     all_sheets[client].to_excel(writer, sheet_name=client, index=False)

# # #             return list(set(updated_clients))

# # #         except Exception as e:
# # #             print(f"[ERROR] Failed to update meetings.xlsx: {e}")
# # #             return []







# # import os
# # import datetime
# # import pandas as pd
# # from typing import List, Dict, Any
# # from difflib import SequenceMatcher
# # from google.oauth2.credentials import Credentials
# # from googleapiclient.discovery import build
# # from dotenv import load_dotenv

# # # Load environment variables
# # load_dotenv()


# # def normalize(name: str) -> str:
# #     """Normalize names for consistent matching."""
# #     return (
# #         name.strip()
# #         .lower()
# #         .replace("_", " ")
# #         .replace("-", " ")
# #         .replace("technologies", "")
# #     )


# # def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.6) -> str:
# #     """Find the best fuzzy match for target among candidates (normalized)."""
# #     best_match, score = None, 0
# #     for c in candidates:
# #         ratio = SequenceMatcher(None, target.lower(), c.lower()).ratio()
# #         if ratio > score:
# #             best_match, score = c, ratio
# #     return best_match if score >= threshold else None


# # def find_client_from_event(title: str, client_sheets: List[str]) -> str:
# #     """Map a calendar event title to the actual client sheet name."""
# #     norm_title = normalize(title)

# #     # Direct containment or exact normalized match
# #     for sheet in client_sheets:
# #         if normalize(sheet) in norm_title or normalize(sheet) == norm_title:
# #             return sheet

# #     # Fallback: fuzzy match
# #     matched = fuzzy_match(norm_title, [normalize(s) for s in client_sheets])
# #     if matched:
# #         for sheet in client_sheets:
# #             if normalize(sheet) == matched:
# #                 return sheet
# #     return None


# # class CalendarHandler:
# #     """
# #     Utility to fetch upcoming meetings from Google Calendar 
# #     and sync them into meetings.xlsx client sheets.
# #     """

# #     def __init__(self, meetings_file: str = "meetings.xlsx", crm_file: str = "crm_data.xlsx"):
# #         self.meetings_file = meetings_file
# #         self.crm_file = crm_file
# #         self.service = self.authenticate()

# #     def authenticate(self):
# #         """Authenticate with Google Calendar using .env credentials."""
# #         creds = Credentials(
# #             token=None,
# #             refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
# #             token_uri="https://oauth2.googleapis.com/token",
# #             client_id=os.getenv("GOOGLE_CLIENT_ID"),
# #             client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
# #             scopes=["https://www.googleapis.com/auth/calendar.readonly"],
# #         )
# #         return build("calendar", "v3", credentials=creds)

# #     def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
# #         """Fetch upcoming events within the next N days."""
# #         now = datetime.datetime.utcnow().isoformat() + "Z"
# #         future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

# #         events_result = (
# #             self.service.events()
# #             .list(
# #                 calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
# #                 timeMin=now,
# #                 timeMax=future,
# #                 singleEvents=True,
# #                 orderBy="startTime",
# #             )
# #             .execute()
# #         )

# #         events = events_result.get("items", [])
# #         meetings = []

# #         for event in events:
# #             title = event.get("summary", "Untitled Meeting")
# #             start = event["start"].get("dateTime", event["start"].get("date"))

# #             attendees = []
# #             if "attendees" in event:
# #                 for att in event["attendees"]:
# #                     name = att.get("displayName") or att.get("email", "Unknown")
# #                     attendees.append(name)

# #             meetings.append({
# #                 "title": title,
# #                 "date": start,
# #                 "participants": attendees
# #             })

# #         return meetings

# #     def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
# #         """
# #         Add fetched events into meetings.xlsx.
# #         - Auto-increments Meeting_IDs per client.
# #         - Looks up Deal_ID from crm_data.xlsx.
# #         Returns: list of updated client sheet names.
# #         """
# #         updated_clients = []
# #         try:
# #             # Load CRM
# #             crm_df = pd.read_excel(self.crm_file)
# #             crm_df["norm"] = crm_df["Client"].str.lower().str.replace("-", " ").str.replace("_", " ")

# #             # Load meetings.xlsx sheets
# #             xls = pd.ExcelFile(self.meetings_file)
# #             client_sheets = xls.sheet_names
# #             all_sheets = {sheet: pd.read_excel(self.meetings_file, sheet_name=sheet) for sheet in client_sheets}

# #             for event in events:
# #                 title = event["title"]

# #                 client_name = find_client_from_event(title, client_sheets)
# #                 if not client_name:
# #                     print(f"[WARN] No client sheet matched for event '{title}'")
# #                     continue

# #                 df = all_sheets[client_name]

# #                 # --- Auto-increment Meeting_ID ---
# #                 if "Meeting_ID" in df.columns and not df.empty:
# #                     last_id = df["Meeting_ID"].dropna().iloc[-1]
# #                     try:
# #                         next_id_num = int(str(last_id).replace("M", "")) + 1
# #                     except ValueError:
# #                         next_id_num = len(df) + 1
# #                     meeting_id = f"M{next_id_num:03d}"
# #                 else:
# #                     meeting_id = "M001"

# #                 # --- Fetch Deal_ID from CRM ---
# #                 crm_match = crm_df[crm_df["norm"] == normalize(client_name)]
# #                 deal_id = crm_match["Deal_ID"].iloc[0] if not crm_match.empty else ""

# #                 # --- Role Detection (simple rule) ---
# #                 reps, clients = [], []
# #                 for p in event["participants"]:
# #                     if "aakashgupta" in p.lower():  # your rep name
# #                         reps.append(p)
# #                     else:
# #                         clients.append(p)

# #                 new_row = {
# #                     "Meeting_ID": meeting_id,
# #                     "Deal_ID": deal_id,
# #                     "Title": event["title"],
# #                     "Date": event["date"],
# #                     "Participants": f"Sales Reps: {', '.join(reps)} | Clients: {', '.join(clients)}",
# #                     "Status": "upcoming",
# #                     "Transcript_File": "",
# #                     "Notes": "Imported from Google Calendar"
# #                 }

# #                 print(f"[DEBUG] Adding new meeting row to '{client_name}': {new_row}")
# #                 df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
# #                 all_sheets[client_name] = df
# #                 updated_clients.append(client_name)
# #                 print(f"✅ Added new meeting for {client_name}: {title} -> {meeting_id}")

# #             # --- Write back ONLY updated sheets ---
# #             if updated_clients:
# #                 with pd.ExcelWriter(self.meetings_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
# #                     for client in set(updated_clients):
# #                         all_sheets[client].to_excel(writer, sheet_name=client, index=False)

# #             return list(set(updated_clients))

# #         except Exception as e:
# #             print(f"[ERROR] Failed to update meetings.xlsx: {e}")
# #             return []






# import os
# import datetime
# from typing import List, Dict, Any
# from google.oauth2.credentials import Credentials
# from googleapiclient.discovery import build
# from dotenv import load_dotenv

# import pandas as pd
# from openpyxl import load_workbook, Workbook

# # Load environment variables
# load_dotenv()


# def normalize(s: str) -> str:
#     if s is None:
#         return ""
#     return str(s).strip().lower().replace("_", " ").replace("-", " ").replace("technologies", "").strip()


# class CalendarHandler:
#     """
#     Minimal, robust calendar -> meetings.xlsx writer using openpyxl.
#     - Keeps logic simple: find sheet, append a row, save.
#     """

#     def __init__(self, meetings_file: str = "meetings.xlsx", crm_file: str = "crm_data.xlsx"):
#         self.meetings_file = meetings_file
#         self.crm_file = crm_file
#         # Only create the Google service if you actually call fetch_upcoming_events
#         self.service = None

#     def authenticate(self):
#         """Authenticate with Google Calendar using .env credentials."""
#         if self.service is not None:
#             return self.service
#         creds = Credentials(
#             token=None,
#             refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
#             token_uri="https://oauth2.googleapis.com/token",
#             client_id=os.getenv("GOOGLE_CLIENT_ID"),
#             client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
#             scopes=["https://www.googleapis.com/auth/calendar.readonly"],
#         )
#         self.service = build("calendar", "v3", credentials=creds)
#         return self.service

#     def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
#         """Fetch upcoming events within the next N days (uses Google Calendar API)."""
#         svc = self.authenticate()
#         now = datetime.datetime.utcnow().isoformat() + "Z"
#         future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

#         events_result = (
#             svc.events()
#             .list(
#                 calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
#                 timeMin=now,
#                 timeMax=future,
#                 singleEvents=True,
#                 orderBy="startTime",
#             )
#             .execute()
#         )

#         events = events_result.get("items", [])
#         meetings = []

#         for event in events:
#             title = event.get("summary", "Untitled Meeting")
#             start = event["start"].get("dateTime", event["start"].get("date"))
#             attendees = []
#             if "attendees" in event:
#                 for att in event["attendees"]:
#                     name = att.get("displayName") or att.get("email") or "Unknown"
#                     attendees.append(name)
#             meetings.append({"title": title, "date": start, "participants": attendees})
#         return meetings

#     def _ensure_workbook(self):
#         """Load existing workbook or create a new one."""
#         if os.path.exists(self.meetings_file):
#             wb = load_workbook(self.meetings_file)
#         else:
#             # create minimal workbook - user can add sheets manually or we try to create from CRM
#             wb = Workbook()
#             # remove default sheet to avoid confusion if empty
#             default = wb.active
#             wb.remove(default)
#             # If CRM available, create sheets for clients
#             if os.path.exists(self.crm_file):
#                 try:
#                     crm_df = pd.read_excel(self.crm_file)
#                     if "Client" in crm_df.columns:
#                         for client in crm_df["Client"].dropna().unique():
#                             ws = wb.create_sheet(title=str(client))
#                 except Exception:
#                     # ignore CRM read errors
#                     pass
#             # make at least one sheet
#             if not wb.sheetnames:
#                 wb.create_sheet(title="Default")
#             wb.save(self.meetings_file)
#         return wb

#     def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
#         """
#         Write events back to meetings.xlsx (one row per event).
#         Returns list of client sheets updated.
#         """
#         updated = []
#         if not events:
#             print("[INFO] No events to add.")
#             return updated

#         # Load CRM for Deal_ID lookup (if available)
#         crm_df = None
#         if os.path.exists(self.crm_file):
#             try:
#                 crm_df = pd.read_excel(self.crm_file, dtype=str)
#                 if "Client" in crm_df.columns:
#                     crm_df["norm_client"] = crm_df["Client"].astype(str).apply(normalize)
#             except Exception as e:
#                 print(f"[WARN] Could not read CRM file: {e}")
#                 crm_df = None

#         # Ensure workbook exists
#         try:
#             wb = self._ensure_workbook()
#         except Exception as e:
#             print(f"[ERROR] Could not open or create {self.meetings_file}: {e}")
#             return updated

#         sheet_names = wb.sheetnames

#         for ev in events:
#             title = str(ev.get("title", "")).strip()
#             date_val = ev.get("date", "")
#             participants = ev.get("participants", []) or []

#             # Find sheet: prefer exact match, then normalized containment, else skip
#             target_sheet = None
#             # exact match
#             for s in sheet_names:
#                 if s.strip().lower() == title.lower():
#                     target_sheet = s
#                     break
#             # containment / normalized match
#             if not target_sheet:
#                 norm_title = normalize(title)
#                 for s in sheet_names:
#                     if normalize(s) in norm_title or norm_title in normalize(s):
#                         target_sheet = s
#                         break

#             if not target_sheet:
#                 print(f"[WARN] No client sheet matched for event '{title}'. Skipping (sheet names: {sheet_names})")
#                 continue

#             ws = wb[target_sheet]

#             # Ensure headers exist in first row (if sheet empty)
#             header_row = 1
#             headers = []
#             if ws.max_row == 0 or all(ws.cell(row=1, column=c).value is None for c in range(1, ws.max_column + 1)):
#                 headers = ["Meeting_ID", "Deal_ID", "Title", "Date", "Participants", "Status", "Transcript_File", "Notes"]
#                 ws.append(headers)
#             else:
#                 headers = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
#                 # if required headers missing, extend headers
#                 required = ["Meeting_ID", "Deal_ID", "Title", "Date", "Participants", "Status", "Transcript_File", "Notes"]
#                 if not all(h in headers for h in required):
#                     # append any missing header columns
#                     for h in required:
#                         if h not in headers:
#                             headers.append(h)
#                     # rewrite first row with new headers (safer approach)
#                     for idx, h in enumerate(headers, start=1):
#                         ws.cell(row=1, column=idx, value=h)

#             # Determine next Meeting_ID by scanning existing Meeting_ID column (if present)
#             try:
#                 midx = headers.index("Meeting_ID") + 1
#             except ValueError:
#                 midx = 1

#             last_num = 0
#             for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=midx, max_col=midx):
#                 val = row[0].value
#                 if not val:
#                     continue
#                 sval = str(val).strip()
#                 if sval.startswith("M"):
#                     try:
#                         n = int(sval[1:])
#                         if n > last_num:
#                             last_num = n
#                     except Exception:
#                         continue
#                 else:
#                     try:
#                         n = int(sval)
#                         if n > last_num:
#                             last_num = n
#                     except Exception:
#                         continue
#             next_num = last_num + 1
#             meeting_id = f"M{next_num:03d}"

#             # Lookup Deal_ID from CRM if available (match by normalized sheet name)
#             deal_id = ""
#             if crm_df is not None:
#                 try:
#                     norm_sheet = normalize(target_sheet)
#                     match = crm_df[crm_df["norm_client"] == norm_sheet]
#                     if not match.empty and "Deal_ID" in match.columns:
#                         deal_id = str(match.iloc[0]["Deal_ID"])
#                 except Exception:
#                     deal_id = ""

#             # Build participants string
#             reps, clients = [], []
#             for p in participants:
#                 if isinstance(p, str) and "aakashgupta" in p.lower():
#                     reps.append(p)
#                 else:
#                     clients.append(p)
#             participant_str = f"Sales Reps: {', '.join(reps)} | Clients: {', '.join(clients)}"

#             # Create new row in the same order as headers
#             new_row_map = {
#                 "Meeting_ID": meeting_id,
#                 "Deal_ID": deal_id,
#                 "Title": title,
#                 "Date": date_val,
#                 "Participants": participant_str,
#                 "Status": "upcoming",
#                 "Transcript_File": "",
#                 "Notes": "Imported from Google Calendar"
#             }
#             row_values = [new_row_map.get(h, "") for h in headers]
#             ws.append(row_values)

#             updated.append(target_sheet)
#             print(f"[DEBUG] Appended to sheet '{target_sheet}': {new_row_map}")

#         # Save workbook
#         try:
#             wb.save(self.meetings_file)
#             print(f"[INFO] Successfully saved updates to {self.meetings_file} (sheets updated: {sorted(set(updated))})")
#         except PermissionError as e:
#             print(f"[ERROR] Permission denied saving {self.meetings_file}: {e}. Please close the file if open in Excel.")
#         except Exception as e:
#             print(f"[ERROR] Failed to save {self.meetings_file}: {e}")

#         return sorted(set(updated))








import os
import datetime
import pandas as pd
from typing import List, Dict, Any
from difflib import SequenceMatcher
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv

from openpyxl import load_workbook, Workbook

# Load environment variables
load_dotenv()


def normalize(s: str) -> str:
    if s is None:
        return ""
    return str(s).strip().lower().replace("_", " ").replace("-", " ").replace("technologies", "").strip()


def format_name(raw: str) -> str:
    """Format a displayName/email into a neat Title Case name."""
    if not raw:
        return "Unknown"
    if "@" in raw:  # if it's an email
        base = raw.split("@")[0]
    else:
        base = raw
    base = base.replace(".", " ").replace("_", " ")
    return " ".join([w.capitalize() for w in base.split()])


class CalendarHandler:
    """
    Calendar handler that fetches Google Calendar events
    and writes them into meetings.xlsx, cleanly formatting participants.
    """

    def __init__(self, meetings_file: str = "meetings.xlsx", crm_file: str = "crm_data.xlsx"):
        self.meetings_file = meetings_file
        self.crm_file = crm_file
        self.service = None

    def authenticate(self):
        """Authenticate with Google Calendar using .env credentials."""
        if self.service:
            return self.service
        creds = Credentials(
            token=None,
            refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=os.getenv("GOOGLE_CLIENT_ID"),
            client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
            scopes=["https://www.googleapis.com/auth/calendar.readonly"],
        )
        self.service = build("calendar", "v3", credentials=creds)
        return self.service

    def fetch_upcoming_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
        """Fetch upcoming events within the next N days."""
        svc = self.authenticate()
        now = datetime.datetime.utcnow().isoformat() + "Z"
        future = (datetime.datetime.utcnow() + datetime.timedelta(days=days_ahead)).isoformat() + "Z"

        events_result = (
            svc.events()
            .list(
                calendarId=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
                timeMin=now,
                timeMax=future,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )

        events = events_result.get("items", [])
        meetings = []

        for event in events:
            title = event.get("summary", "Untitled Meeting")
            start = event["start"].get("dateTime", event["start"].get("date"))
            attendees = []
            if "attendees" in event:
                for att in event["attendees"]:
                    name = att.get("displayName") or att.get("email", "Unknown")
                    attendees.append(name)

            meetings.append({"title": title, "date": start, "participants": attendees})
        return meetings

    def _ensure_workbook(self):
        """Load existing workbook or create new one with sheets for CRM clients."""
        if os.path.exists(self.meetings_file):
            wb = load_workbook(self.meetings_file)
        else:
            wb = Workbook()
            default = wb.active
            wb.remove(default)
            if os.path.exists(self.crm_file):
                crm_df = pd.read_excel(self.crm_file)
                if "Client" in crm_df.columns:
                    for client in crm_df["Client"].dropna().unique():
                        wb.create_sheet(title=str(client))
            if not wb.sheetnames:
                wb.create_sheet(title="Default")
            wb.save(self.meetings_file)
        return wb

    def add_to_meetings_excel(self, events: List[Dict[str, Any]]) -> List[str]:
        """Add fetched events into meetings.xlsx (appends as new rows)."""
        updated = []
        if not events:
            return updated

        # Load CRM for Deal_ID lookup
        crm_df = None
        if os.path.exists(self.crm_file):
            crm_df = pd.read_excel(self.crm_file, dtype=str)
            if "Client" in crm_df.columns:
                crm_df["norm_client"] = crm_df["Client"].astype(str).apply(normalize)

        wb = self._ensure_workbook()
        sheet_names = wb.sheetnames

        for ev in events:
            title = str(ev.get("title", "")).strip()
            date_val = ev.get("date", "")
            participants = ev.get("participants", []) or []

            # --- Find matching client sheet ---
            target_sheet = None
            for s in sheet_names:
                if s.strip().lower() == title.lower():
                    target_sheet = s
                    break
            if not target_sheet:
                norm_title = normalize(title)
                for s in sheet_names:
                    if normalize(s) in norm_title or norm_title in normalize(s):
                        target_sheet = s
                        break
            if not target_sheet:
                print(f"[WARN] No client sheet matched for event '{title}'. Skipping.")
                continue

            ws = wb[target_sheet]

            # Ensure headers
            headers = ["Meeting_ID", "Deal_ID", "Title", "Date", "Participants", "Status", "Transcript_File", "Notes"]
            if ws.max_row == 0:
                ws.append(headers)
            else:
                current_headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
                if not all(h in current_headers for h in headers):
                    ws.delete_rows(1)
                    ws.append(headers)

            # Auto-increment Meeting_ID
            existing_ids = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
            nums = [int(str(mid).replace("M", "")) for mid in existing_ids if mid and str(mid).startswith("M")]
            next_id = f"M{(max(nums) + 1) if nums else 1:03d}"

            # Deal_ID from CRM
            deal_id = ""
            if crm_df is not None:
                match = crm_df[crm_df["norm_client"] == normalize(target_sheet)]
                if not match.empty and "Deal_ID" in match.columns:
                    deal_id = str(match.iloc[0]["Deal_ID"])

            # Format participants (Name + Role)
            formatted = []
            for p in participants:
                name = format_name(p)
                if "aakashgupta" in str(p).lower():
                    formatted.append(f"{name} (Dice)")
                else:
                    formatted.append(f"{name} (Client)")
            participants_str = ", ".join(formatted)

            new_row = [next_id, deal_id, title, date_val, participants_str, "upcoming", "", "Imported from Google Calendar"]
            ws.append(new_row)
            updated.append(target_sheet)

            print(f"✅ Added meeting to {target_sheet}: {new_row}")

        try:
            wb.save(self.meetings_file)
            print(f"[INFO] Saved updates to {self.meetings_file}")
        except PermissionError:
            print(f"[ERROR] Permission denied. Please close {self.meetings_file} in Excel before running again.")

        return updated
