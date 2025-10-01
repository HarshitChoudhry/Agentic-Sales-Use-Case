# # import os
# # import pandas as pd
# # from utils.excel_handler import ExcelHandler
# # from utils.calendar_handler import CalendarHandler
# # from agents.meeting_prep import MeetingPreparationAgent
# # from agents.faq_agent import FAQAgent
# # from agents.coaching_agent import CoachingAgent
# # from agents.deal_exec import DealExecutionAgent


# # def sync_calendar_to_meetings(meetings_file="meetings.xlsx", crm_file="crm_data.xlsx"):
# #     """
# #     Sync upcoming meetings from Google Calendar into meetings.xlsx client sheets.
# #     """
# #     print("\n--- Syncing Google Calendar with meetings.xlsx ---")
# #     cal = CalendarHandler(meetings_file)
# #     events = cal.fetch_upcoming_events(days_ahead=7)

# #     if not events:
# #         print("⚠️ No upcoming events found in Google Calendar.")
# #         return

# #     # Load CRM to map Account_Name -> Deal_ID
# #     crm_df = pd.read_excel(crm_file, sheet_name="CRM_Data")

# #     for _, row in crm_df.iterrows():
# #         account = row["Account_Name"]
# #         deal_id = row["Deal_ID"]

# #         # Match events whose title contains client name (simple heuristic)
# #         matched_events = [e for e in events if account.lower() in e["title"].lower()]

# #         if matched_events:
# #             cal.add_to_meetings_excel(account.replace(" ", "_"), deal_id, matched_events)


# # def run_pipeline(crm_file="crm_data.xlsx", meetings_file="meetings.xlsx"):
# #     """
# #     Orchestrate the agent workflow for all meetings across clients.
# #     """
# #     # Step 1: Sync Calendar -> Meetings Excel
# #     sync_calendar_to_meetings(meetings_file, crm_file)

# #     handler = ExcelHandler(crm_file, meetings_file)

# #     # Step 2: Load all client sheets
# #     xls = pd.ExcelFile(meetings_file)
# #     clients = xls.sheet_names

# #     # Step 3: Initialize agents
# #     prep_agent = MeetingPreparationAgent(crm_file, meetings_file)
# #     faq_agent = FAQAgent(crm_file, meetings_file)
# #     coach_agent = CoachingAgent(crm_file, meetings_file)
# #     deal_agent = DealExecutionAgent(crm_file, meetings_file)

# #     for client in clients:
# #         print(f"\n--- Processing Client: {client} ---")
# #         df = pd.read_excel(meetings_file, sheet_name=client)

# #         for _, row in df.iterrows():
# #             meeting_id = row["Meeting_ID"]
# #             status = str(row["Status"]).lower()

# #             print(f"Processing {meeting_id} ({status})...")

# #             if status == "upcoming":
# #                 # Run Meeting Preparation
# #                 try:
# #                     prep_output = prep_agent.prepare_meeting(meeting_id, client)
# #                     print(f"✅ Meeting Prep done for {meeting_id}")
# #                 except Exception as e:
# #                     print(f"⚠️ Meeting Prep failed for {meeting_id}: {e}")

# #             elif status == "past":
# #                 # Run FAQ Agent
# #                 try:
# #                     faq_output = faq_agent.process(meeting_id, client)
# #                     print(f"✅ FAQ extraction done for {meeting_id}, {len(faq_output)} Q&A found")
# #                 except Exception as e:
# #                     print(f"⚠️ FAQ Agent failed for {meeting_id}: {e}")

# #                 # Run Coaching Agent
# #                 try:
# #                     coach_output = coach_agent.process(meeting_id, client)
# #                     print(f"✅ Coaching report done for {meeting_id}")
# #                 except Exception as e:
# #                     print(f"⚠️ Coaching Agent failed for {meeting_id}: {e}")

# #                 # Run Deal Execution Agent
# #                 try:
# #                     deal_output = deal_agent.process(meeting_id, client)
# #                     print(f"✅ Deal Execution summary done for {meeting_id}")
# #                 except Exception as e:
# #                     print(f"⚠️ Deal Execution Agent failed for {meeting_id}: {e}")

# #     print("\nPipeline completed for all clients!")


# # if __name__ == "__main__":
# #     run_pipeline()







# import os
# import pandas as pd
# from agents.meeting_prep import MeetingPreparationAgent
# from agents.faq_agent import FAQAgent
# from agents.coaching_agent import CoachingAgent
# from agents.deal_exec import DealExecutionAgent
# from utils.excel_handler import ExcelHandler
# from utils.calendar_handler import CalendarHandler


# def process_client(client_name: str):
#     """Run all agents for a specific client."""
#     excel_handler = ExcelHandler("crm_data.xlsx", "meetings.xlsx")
#     meetings = excel_handler.list_meetings(client_name)

#     meeting_prep = MeetingPreparationAgent()
#     faq_agent = FAQAgent()
#     coaching_agent = CoachingAgent()
#     deal_exec_agent = DealExecutionAgent()

#     for m in meetings:
#         meeting_id = m["Meeting_ID"]
#         status = m["Status"]

#         try:
#             if status == "upcoming":
#                 print(f"Running Meeting Prep for {client_name} - {meeting_id}")
#                 meeting_prep.prepare_meeting(meeting_id, client_name)

#             elif status == "past":
#                 print(f"Running FAQ Agent for {client_name} - {meeting_id}")
#                 faq_agent.process(meeting_id, client_name)

#                 print(f"Running Coaching Agent for {client_name} - {meeting_id}")
#                 coaching_agent.process(meeting_id, client_name)

#                 print(f"Running Deal Execution Agent for {client_name} - {meeting_id}")
#                 deal_exec_agent.process(meeting_id, client_name)

#         except Exception as e:
#             print(f"⚠️ Agent failed for {client_name} {meeting_id}: {e}")


# def run_pipeline(selected_clients=None):
#     """Run pipeline only for selected clients (or all if none specified)."""
#     xls = pd.ExcelFile("meetings.xlsx")
#     all_clients = xls.sheet_names
#     clients_to_process = selected_clients or all_clients

#     for client in clients_to_process:
#         print(f"\n--- Processing Client: {client} ---")
#         process_client(client)


# if __name__ == "__main__":
#     # Sync calendar first
#     calendar_handler = CalendarHandler("meetings.xlsx")
#     events = calendar_handler.fetch_upcoming_events(days_ahead=7)
#     updated_clients = calendar_handler.add_to_meetings_excel(events)

#     if not updated_clients:
#         print("No new calendar meetings found. Processing all clients.")
#         run_pipeline()
#     else:
#         print(f"Processing only updated clients from calendar: {updated_clients}")
#         run_pipeline(updated_clients)





import pandas as pd
from agents.meeting_prep import MeetingPreparationAgent
from agents.faq_agent import FAQAgent
from agents.coaching_agent import CoachingAgent
from agents.deal_exec import DealExecutionAgent
from utils.excel_handler import ExcelHandler
from utils.calendar_handler import CalendarHandler


def process_client(client_name: str):
    """Run all agents for a specific client."""
    excel_handler = ExcelHandler("crm_data.xlsx", "meetings.xlsx")
    meetings = excel_handler.list_meetings(client_name)

    meeting_prep = MeetingPreparationAgent()
    faq_agent = FAQAgent()
    coaching_agent = CoachingAgent()
    deal_exec_agent = DealExecutionAgent()

    for m in meetings:
        meeting_id = m["Meeting_ID"]
        status = m["Status"]

        try:
            if status == "upcoming":
                print(f"Running Meeting Prep for {client_name} - {meeting_id}")
                try:
                    meeting_prep.prepare_meeting(meeting_id, client_name)
                except Exception as e:
                    print(f"⚠️ Meeting Prep failed for {client_name} {meeting_id}: {e}")

            elif status == "past":
                print(f"Running FAQ Agent for {client_name} - {meeting_id}")
                try:
                    faq_results = faq_agent.process(meeting_id, client_name)
                    if not faq_results:
                        print(f"[INFO] No FAQ questions found for {client_name} {meeting_id}")
                except Exception as e:
                    print(f"⚠️ FAQ Agent failed for {client_name} {meeting_id}: {e}")

                print(f"Running Coaching Agent for {client_name} - {meeting_id}")
                try:
                    coaching_agent.process(meeting_id, client_name)
                except Exception as e:
                    print(f"⚠️ Coaching Agent failed for {client_name} {meeting_id}: {e}")

                print(f"Running Deal Execution Agent for {client_name} - {meeting_id}")
                try:
                    deal_exec_agent.process(meeting_id, client_name)
                except Exception as e:
                    print(f"⚠️ Deal Execution Agent failed for {client_name} {meeting_id}: {e}")

        except Exception as e:
            print(f"⚠️ Agent pipeline failed for {client_name} {meeting_id}: {e}")

    # --- Account-level summary ---
    try:
        print(f"Running Account-Level Deal Execution Summary for {client_name}")
        deal_exec_agent.process_account(client_name)
    except Exception as e:
        print(f"⚠️ Account-Level Summary failed for {client_name}: {e}")


def run_pipeline(selected_clients=None):
    """Run pipeline only for selected clients (or all if none specified)."""
    xls = pd.ExcelFile("meetings.xlsx")
    all_clients = xls.sheet_names
    clients_to_process = selected_clients or all_clients

    for client in clients_to_process:
        print(f"\n--- Processing Client: {client} ---")
        process_client(client)


if __name__ == "__main__":
    # Sync calendar first
    calendar_handler = CalendarHandler("meetings.xlsx")
    events = calendar_handler.fetch_upcoming_events(days_ahead=3)
    updated_clients = calendar_handler.add_to_meetings_excel(events)

    if not updated_clients:
        print("No new calendar meetings found. Processing all clients.")
        run_pipeline()
    else:
        print(f"Processing only updated clients from calendar: {updated_clients}")
        run_pipeline(updated_clients)
