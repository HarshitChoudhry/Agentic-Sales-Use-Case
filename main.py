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
        status = str(m.get("Status", "")).lower()

        try:
            if status == "upcoming":
                # --- Run Meeting Preparation ---
                print(f"Running Meeting Prep for {client_name} - {meeting_id}")
                try:
                    meeting_prep.prepare_meeting(meeting_id, client_name)
                    print(f"✅ Meeting Prep completed for {meeting_id}")
                except Exception as e:
                    print(f"⚠️ Meeting Prep failed for {client_name} {meeting_id}: {e}")

            elif status == "past":
                # --- Run FAQ Agent ---
                print(f"Running FAQ Agent for {client_name} - {meeting_id}")
                try:
                    faq_results = faq_agent.process(meeting_id, client_name)
                    if not faq_results:
                        print(f"[INFO] No FAQ questions found for {client_name} {meeting_id}")
                    else:
                        print(f"✅ FAQ extraction completed for {meeting_id}, {len(faq_results)} entries")
                except Exception as e:
                    print(f"⚠️ FAQ Agent failed for {client_name} {meeting_id}: {e}")

                # --- Run Coaching Agent (only works with transcripts -> past meetings) ---
                print(f"Running Coaching Agent for {client_name} - {meeting_id}")
                try:
                    coaching_agent.process(meeting_id, client_name)
                    print(f"✅ Coaching report completed for {meeting_id}")
                except Exception as e:
                    print(f"⚠️ Coaching Agent failed for {client_name} {meeting_id}: {e}")

                # --- Run Deal Execution Agent ---
                print(f"Running Deal Execution Agent for {client_name} - {meeting_id}")
                try:
                    deal_exec_agent.process(meeting_id, client_name)
                    print(f"✅ Deal Execution summary completed for {meeting_id}")
                except Exception as e:
                    print(f"⚠️ Deal Execution Agent failed for {client_name} {meeting_id}: {e}")

        except Exception as e:
            print(f"⚠️ Agent pipeline failed for {client_name} {meeting_id}: {e}")

    # --- Account-level summary ---
    try:
        print(f"Running Account-Level Deal Execution Summary for {client_name}")
        deal_exec_agent.process_account(client_name)
        print(f"✅ Account-Level Summary completed for {client_name}")
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
    # --- Step 1: Sync calendar first ---
    calendar_handler = CalendarHandler("meetings.xlsx")
    events = calendar_handler.fetch_upcoming_events(days_ahead=3)
    updated_clients = calendar_handler.add_to_meetings_excel(events)

    # --- Step 2: Run pipeline ---
    if not updated_clients:
        print("No new calendar meetings found. Processing all clients.")
        run_pipeline()
    else:
        print(f"Processing only updated clients from calendar: {updated_clients}")
        run_pipeline(updated_clients)
