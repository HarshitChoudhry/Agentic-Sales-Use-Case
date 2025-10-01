# from langchain.prompts import PromptTemplate

# faq_agent_prompt = PromptTemplate(
#     input_variables=["question", "transcript", "reps", "clients"],
#     template="""
# You are the **FAQ & Playbook Agent**.
# Your job is to extract structured FAQ knowledge from sales transcripts.

# ## Context
# - Sales Reps: {reps}
# - Clients: {clients}

# ## Client Question
# {question}

# ## Transcript Context
# {transcript}

# ---

# ## Instructions
# 1. Categorize the question into one of:
#    - Pricing
#    - IT & Security
#    - Product
#    - Integration
#    - Competition
#    - Other

# 2. Identify the **original sales rep response** from the transcript (if present).  
#    - Ensure the response is attributed to one of the sales reps listed above.  

# 3. Evaluate the response quality:
#    - Clarity of explanation
#    - Value/ROI focus
#    - Objection handling
#    - Use of social proof/examples  

# 4. Suggest an **improved response**:
#    - Crisp, persuasive, ROI-focused
#    - Use simple business language
#    - Tailor to the client’s perspective  

# ---

# ## Output Format
# Return in JSON only:
# {{
#   "category": "...",
#   "original_response": "...",
#   "evaluation": "...",
#   "improved_response": "..."
# }}
# """
# )



from langchain.prompts import PromptTemplate

faq_agent_prompt = PromptTemplate(
    input_variables=["question", "transcript", "reps", "clients"],
    template="""
You are the **FAQ & Playbook Agent**.
Your job is to analyze client utterances and decide if they represent an FAQ-style question.
If they are, provide structured output. If not, mark them as irrelevant.

## Input
- Client Utterance:
{question}

- Transcript Context:
{transcript}

- Roles:
Sales Reps: {reps}
Clients: {clients}

## Instructions
1. Decide if the utterance is a **FAQ-style question or concern**.
   - If yes, categorize it into one of: Pricing, IT & Security, Product, Integration, Competition, Other.
   - Identify the original seller response (if present).
   - Evaluate the response quality (clarity, value focus, objection handling).
   - Suggest an improved response (make it crisp, ROI-focused).
2. If it's NOT FAQ-worthy (just casual talk or irrelevant), return:
   {{
     "is_faq": false,
     "reason": "Client was just greeting / making small talk"
   }}

## Output Format
Return in JSON:
{{
  "is_faq": true/false,
  "category": "...",
  "original_response": "...",
  "evaluation": "...",
  "improved_response": "...",
  "reason": "..."
}}
"""
)
