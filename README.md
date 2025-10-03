# 🧠 Sales Agentic AI System/ Sales Coach

The Sales Agentic AI System is an intelligent assistant that automates the **end-to-end sales meeting workflow**. It helps sales teams prepare effectively, capture FAQs from clients, receive coaching feedback, and generate structured deal execution reports. By leveraging **OpenAI** and **Google Gemini** models, the system transforms raw transcripts and CRM data into actionable business insights.

---

## 🔍 How It Works

1. **Meeting Preparation (Before Meeting)**

   * Reads CRM data (`crm_data.xlsx`) and meeting schedule (`meetings.xlsx`).
   * Generates a **Prep Pack** with stakeholder mapping, agenda, anticipated objections, talking points, and follow-up draft.
   * Output: Excel guidance + PDF prep pack + draft follow-up email.

2. **FAQ Extraction (After Meeting)**

   * Loads transcripts (`transcripts/{client}_{meeting}.json`).
   * Groups client utterances into meaningful blocks.
   * Identifies FAQs, improves both questions and responses, and categorizes them.
   * Output: Excel FAQ knowledge base.

3. **Coaching Report (After Meeting)**

   * Evaluates reps’ skills (selling, closing, objection handling, tone, etc.).
   * Generates a **coaching scorecard** with strengths, improvements, and nudges.
   * Output: Coaching Excel + Coaching PDF + optional radar charts.

4. **Deal Execution Summary (After Meeting)**

   * Produces structured meeting summaries with overview, risks, action items, and decision process.
   * Generates both meeting-level and account-level summaries.
   * Output: Deal execution Excel + PDF summaries.

The system pipeline is orchestrated via `main.py`, which automatically triggers the appropriate agents based on meeting status (upcoming/past).

---

## 💻 How to Run

1. **Clone and Setup Environment**

   ```bash
   git clone https://github.com/HarshitChoudhry/sales-agentic-ai-system.git
   cd sales-agentic-ai-system
   python -m venv myenv
   source myenv/bin/activate   # or myenv\Scripts\activate on Windows
   pip install -r requirements.txt
   ```

2. **Set Up Environment Variables**
   Create a `.env` file in the root directory and add your keys:

   ```
   OPENAI_API_KEY=your_openai_key
   GOOGLE_API_KEY=your_google_key
   ```

3. **Run the Pipeline**

   ```bash
   python main.py
   streamlit run app.py
   ```

4. **View Outputs**

   * `outputs/prep_packs/` → Meeting prep PDFs.
   * `outputs/faq_knowledge.xlsx` → FAQ knowledge base.
   * `outputs/coaching_reports/` → Coaching Excel & PDFs.
   * `outputs/deal_summaries/` → Deal execution summaries.

---

## 🛠️ Tech Stack

* **OpenAI GPT** – for coaching and deal execution analysis.
* **Google Gemini** – for FAQ extraction and improved responses.
* **LangChain** – for LLM orchestration and chaining prompts.
* **Pandas** – Excel handling and reporting.
* **Fpdf/Matplotlib** – PDF report generation and radar charts.
* **Streamlit** – Dashboard visualization (`app.py`).

---



