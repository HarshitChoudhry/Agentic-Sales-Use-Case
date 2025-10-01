import os
import streamlit as st
import pandas as pd
from utils.calendar_handler import CalendarHandler
from agents.meeting_prep import MeetingPreparationAgent
from agents.deal_exec import DealExecutionAgent
from agents.coaching_agent import CoachingAgent
from utils.output_writer import OutputWriter

# -----------------------
# Initialize
# -----------------------
st.set_page_config(page_title="Agentic Sales Assistant", layout="wide")
st.title("🤖 Agentic Sales Assistant Dashboard")

calendar_handler = CalendarHandler()
writer = OutputWriter()

meetings_file = "meetings.xlsx"
crm_file = "crm_data.xlsx"

# -----------------------
# 1. Upcoming Meetings Sync
# -----------------------
st.header("📅 Upcoming Meetings")


if os.path.exists(meetings_file):
    xls = pd.ExcelFile(meetings_file)
    # Collect upcoming meetings
    upcoming_data = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(meetings_file, sheet_name=sheet)
        if "Status" in df.columns:
            ups = df[df["Status"].str.lower() == "upcoming"].copy()
            if not ups.empty:
                ups["Client"] = sheet
                upcoming_data.append(ups)
    if upcoming_data:
        all_upcoming = pd.concat(upcoming_data)
        st.dataframe(all_upcoming[["Meeting_ID", "Client", "Title", "Date", "Participants"]])
    else:
        st.info("✅ No upcoming meetings scheduled in next 7 days.")
else:
    st.warning("meetings.xlsx not found yet.")

# -----------------------
# 2. Generate Reports
# -----------------------
st.header("📝 Generate Reports & Prep Packs")

meeting_choice = None
if os.path.exists(meetings_file):
    df_all = pd.read_excel(meetings_file, None)  # dict of all sheets
    meeting_options = []
    for client, df in df_all.items():
        if "Status" in df.columns:
            for _, row in df[df["Status"].str.lower() == "upcoming"].iterrows():
                meeting_options.append(f"{row['Meeting_ID']} | {client} | {row['Title']}")
    meeting_choice = st.selectbox("Select an upcoming meeting", ["-- Select --"] + meeting_options)

if meeting_choice and meeting_choice != "-- Select --":
    meeting_id, client_name, title = [x.strip() for x in meeting_choice.split("|")]

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📄 Generate Summary Report"):
            agent = DealExecutionAgent(crm_path=crm_file, meetings_path=meetings_file)
            result = agent.process_account(client_name)
            st.success(f"Summary generated for {client_name}")
            summary_file = f"outputs/deal_summaries/{client_name}_account_summary.pdf"
            if os.path.exists(summary_file):
                with open(summary_file, "rb") as f:
                    st.download_button("⬇️ Download Summary Report", f, file_name=os.path.basename(summary_file))

    with col2:
        if st.button("📑 Generate Prep Pack"):
            agent = MeetingPreparationAgent(crm_path=crm_file, meetings_path=meetings_file)
            result = agent.prepare_meeting(meeting_id, client_name)
            st.success(f"Prep Pack generated for {client_name}")
            prep_file = f"outputs/prep_packs/{client_name}_{meeting_id}_prep.pdf"
            if os.path.exists(prep_file):
                with open(prep_file, "rb") as f:
                    st.download_button("⬇️ Download Prep Pack", f, file_name=os.path.basename(prep_file))


# -----------------------
# 3. Coaching Scorecards
# -----------------------
st.header("📊 Coaching Skill Scorecards")

if meeting_choice and meeting_choice != "-- Select --":
    meeting_id, client_name, title = [x.strip() for x in meeting_choice.split("|")]

    if st.button("🎯 Generate Coaching Report"):
        agent = CoachingAgent(crm_path=crm_file, meetings_path=meetings_file)
        report = agent.process(meeting_id, client_name)
        st.success(f"Coaching report generated for {client_name} - {meeting_id}")

        if report.get("Averaged_Scores"):
            df_scores = pd.DataFrame(list(report["Averaged_Scores"].items()), columns=["Skill", "Score"])
            st.subheader("📌 Skill Scorecard")
            st.table(df_scores)

            # Show Radar chart
            radar_file = f"outputs/{client_name}_{meeting_id}_skills.png"
            if os.path.exists(radar_file):
                st.image(radar_file, caption="Skill Radar Chart", use_column_width=True)
