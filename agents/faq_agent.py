import os
import json
from typing import Dict, Any, List
from langchain_google_genai import ChatGoogleGenerativeAI
from prompts.faq_agent_prompt import faq_agent_prompt
from utils.transcript_loader import TranscriptLoader
from utils.transcript_chunker import TranscriptChunker
from utils.excel_handler import ExcelHandler
from utils.output_writer import OutputWriter
from dotenv import load_dotenv

load_dotenv()


class FAQAgent:
    def __init__(
        self,
        crm_path: str = "crm_data.xlsx",
        meetings_path: str = "meetings.xlsx",
        model: str = "gemini-2.5-flash",
        paragraph_threshold: int = 5,
    ):
        """
        Initialize FAQ Agent with Google Gemini.
        
        Args:
            crm_path: Path to CRM Excel file
            meetings_path: Path to meetings Excel file
            model: Gemini model name (default: gemini-2.5-flash)
            paragraph_threshold: Minimum number of sentences to group into a paragraph
        """
        self.transcript_loader = TranscriptLoader(meetings_path)
        self.excel_handler = ExcelHandler(crm_path, meetings_path)
        self.writer = OutputWriter()
        self.paragraph_threshold = paragraph_threshold
        
        # Use gpt-3.5-turbo for chunker tokenization (tiktoken doesn't support Gemini)
        self.chunker = TranscriptChunker(model_name="gpt-3.5-turbo")
        
        # Initialize Google Gemini
        google_api_key = os.getenv("GOOGLE_API_KEY")
        if not google_api_key:
            raise ValueError("GOOGLE_API_KEY not found in environment variables. Please add it to your .env file")
        
        try:
            self.llm = ChatGoogleGenerativeAI(
                model=model,
                temperature=0.3,
                google_api_key=google_api_key,
                convert_system_message_to_human=True  # Better compatibility
            )
            print(f"[INFO] Using Google Generative AI model: {model}")
        except Exception as e:
            raise RuntimeError(f"Failed to initialize Google AI: {e}\n"
                             f"Make sure your API key is valid and has access to the model.")
        
        # Create chain
        self.chain = faq_agent_prompt | self.llm

    def get_roles_from_meeting(self, meeting_id: str, client_name: str):
        """
        Parse rep vs client roles dynamically from meetings.xlsx Participants column.
        Returns (reps, clients) lists of names (strings).
        """
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
        if not meeting_info:
            return [], []
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
                if "aakashgupta" in p.lower():
                    reps.append(p)
                else:
                    clients.append(p)
        return reps, clients

    def normalize_speaker_name(self, speaker: str) -> str:
        """Normalize speaker name by removing extra whitespace and converting to lowercase."""
        return " ".join(speaker.strip().lower().split())

    def is_client_speaker(self, speaker: str, client_names: List[str]) -> bool:
        """
        Check if speaker matches any of the client names.
        Uses fuzzy matching - checks if any part of client name appears in speaker.
        """
        speaker_norm = self.normalize_speaker_name(speaker)
        
        for client in client_names:
            client_norm = self.normalize_speaker_name(client)
            if client_norm in speaker_norm or speaker_norm in client_norm:
                return True
            
            # Also check first name and last name separately
            client_parts = client_norm.split()
            speaker_parts = speaker_norm.split()
            
            if len(client_parts) >= 2 and len(speaker_parts) >= 2:
                if client_parts[0] in speaker_parts and client_parts[-1] in speaker_parts:
                    return True
        
        return False

    def group_client_utterances_into_paragraphs(
        self, 
        dialogue: List[Dict[str, str]], 
        clients: List[str],
        min_utterances: int = 5,
        min_words: int = 20,
        rep_response_threshold: int = 20
    ) -> List[Dict[str, Any]]:
        """
        Group consecutive client utterances into meaningful conversation blocks/paragraphs.
        
        Strategy:
        - Collect client utterances across multiple back-and-forth exchanges
        - Only break when rep gives a substantial, detailed response
        - Filter out trivial paragraphs during collection, not after
        
        Args:
            min_utterances: Minimum client utterances to consider as a paragraph
            min_words: Minimum total words to consider as a paragraph
            rep_response_threshold: Rep response word count that signals end of topic
        
        Returns list of paragraph dicts
        """
        paragraphs = []
        current_paragraph = []
        
        for idx, entry in enumerate(dialogue):
            text = entry.get("sentence", "").strip()
            speaker = entry.get("speaker", "").strip()
            
            if not text or not speaker:
                continue
            
            is_client = self.is_client_speaker(speaker, clients)
            
            if is_client:
                # Accumulate client utterances
                current_paragraph.append({
                    "text": text,
                    "speaker": speaker,
                    "idx": idx,
                    "timestamp": entry.get("timestamp", "")
                })
            else:
                # Rep is speaking - check if it's a substantial/detailed response
                word_count = len(text.split())
                
                # Only finalize if:
                # 1. Rep gives substantial response (detailed answer)
                # 2. We have accumulated enough client content
                if word_count >= rep_response_threshold and current_paragraph:
                    finalized = self._finalize_paragraph(current_paragraph)
                    
                    # Apply filters immediately
                    if (finalized and 
                        finalized['utterance_count'] >= min_utterances and
                        len(finalized['text'].split()) >= min_words and
                        self._contains_meaningful_content(finalized['text'])):
                        paragraphs.append(finalized)
                    
                    current_paragraph = []
                # Brief acknowledgments don't break the conversation flow
        
        # Handle last paragraph
        if current_paragraph:
            finalized = self._finalize_paragraph(current_paragraph)
            if (finalized and 
                finalized['utterance_count'] >= min_utterances and
                len(finalized['text'].split()) >= min_words and
                self._contains_meaningful_content(finalized['text'])):
                paragraphs.append(finalized)
        
        return paragraphs

    def _contains_meaningful_content(self, text: str) -> bool:
        """
        Check if text contains meaningful content beyond greetings/fillers.
        """
        text_lower = text.lower()
        
        # Common filler words/phrases
        fillers = {
            'hello', 'hi', 'hey', 'okay', 'ok', 'yeah', 'yes', 'no',
            'thanks', 'thank', 'sure', 'alright', 'got', 'it', 'good',
            'great', 'fine', 'right', 'well', 'um', 'uh', 'hmm'
        }
        
        words = [w.strip('.,!?') for w in text_lower.split()]
        meaningful_words = [w for w in words if w not in fillers and len(w) > 2]
        
        # Need at least 10 meaningful words
        return len(meaningful_words) >= 10

    def _is_substantial_paragraph(
        self, 
        paragraph: Dict[str, Any], 
        min_utterances: int, 
        min_words: int
    ) -> bool:
        """
        Check if a paragraph is substantial enough to process.
        NOTE: This is now mostly redundant as filtering happens during collection.
        """
        if not paragraph:
            return False
        
        if paragraph['utterance_count'] < min_utterances:
            return False
        
        word_count = len(paragraph['text'].split())
        if word_count < min_words:
            return False
        
        return self._contains_meaningful_content(paragraph['text'])

    def _finalize_paragraph(self, utterances: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Convert a list of utterances into a finalized paragraph dict."""
        if not utterances:
            return None
        
        combined_text = " ".join([u["text"] for u in utterances])
        speaker = utterances[0]["speaker"]
        
        return {
            "text": combined_text,
            "speaker": speaker,
            "start_idx": utterances[0]["idx"],
            "end_idx": utterances[-1]["idx"],
            "utterance_count": len(utterances),
            "first_timestamp": utterances[0].get("timestamp", ""),
            "last_timestamp": utterances[-1].get("timestamp", "")
        }

    def get_surrounding_context(
        self, 
        dialogue: List[Dict[str, str]], 
        start_idx: int, 
        end_idx: int,
        context_window: int = 5
    ) -> str:
        """
        Get surrounding conversation context for a paragraph.
        Returns formatted transcript with context before and after.
        """
        # Get context before and after
        context_start = max(0, start_idx - context_window)
        context_end = min(len(dialogue), end_idx + context_window + 1)
        
        lines = []
        for entry in dialogue[context_start:context_end]:
            speaker = entry.get("speaker", "Unknown")
            text = entry.get("sentence", "")
            timestamp = entry.get("timestamp", "")
            
            if timestamp:
                lines.append(f"{speaker} [{timestamp}]: {text}")
            else:
                lines.append(f"{speaker}: {text}")
        
        return "\n".join(lines)

    def process(self, meeting_id: str, client_name: str) -> List[Dict[str, Any]]:
        """
        Process FAQ extraction for a past meeting using paragraph-based approach.
        Returns list of the structured rows saved to FAQ workbook.
        """
        # --- Step 1: Load transcript ---
        try:
            dialogue, readable = self.transcript_loader.load_transcript(
                meeting_id, client_name
            )
            print(f"[INFO] Loaded transcript with {len(dialogue)} entries.")
        except FileNotFoundError as e:
            print(f"[WARN] Transcript not found for {meeting_id} / {client_name}: {e}")
            return []
        except Exception as e:
            print(f"[ERROR] Failed loading transcript for {meeting_id} / {client_name}: {e}")
            return []

        # --- Step 2: Detect roles dynamically ---
        reps, clients = self.get_roles_from_meeting(meeting_id, client_name)
        print(f"[INFO] Detected reps: {reps}, clients: {clients}")
        
        # Get meeting metadata
        meeting_info = self.excel_handler.get_meeting_info(meeting_id, client_name)
        meeting_date = meeting_info.get("Date", "") if meeting_info else ""

        # --- Step 3: Group client utterances into paragraphs with aggressive filtering ---
        all_paragraphs = self.group_client_utterances_into_paragraphs(
            dialogue, 
            clients,
            min_utterances=5,        # At least 5 client utterances
            min_words=20,            # At least 20 words total
            rep_response_threshold=20 # Rep must give 20+ word response to break
        )
        
        if not all_paragraphs:
            print(f"[INFO] No substantial client paragraphs detected for {client_name} in {meeting_id}")
            print(f"[HINT] Try lowering min_utterances or min_words thresholds if this seems wrong")
            return []
        
        print(f"[INFO] Created {len(all_paragraphs)} substantial conversation blocks (filtered from raw transcript)")
        print(f"[DEBUG] Sample paragraphs:")
        for i, p in enumerate(all_paragraphs[:3]):
            preview = p['text'][:150] + "..." if len(p['text']) > 150 else p['text']
            word_count = len(p['text'].split())
            print(f"  [{i+1}] {p['utterance_count']} utterances, {word_count} words: {preview}")
        
        # Use filtered paragraphs
        paragraphs = all_paragraphs

        results = []
        
        # --- Step 4: Process each paragraph with LLM ---
        for idx, paragraph in enumerate(paragraphs):
            paragraph_text = paragraph['text']
            speaker = paragraph['speaker']
            
            # Get surrounding context for better understanding
            context = self.get_surrounding_context(
                dialogue, 
                paragraph['start_idx'], 
                paragraph['end_idx'],
                context_window=3
            )
            
            # Call LLM to analyze this paragraph
            try:
                response = self.chain.invoke({
                    "question": f"{speaker}: {paragraph_text}",
                    "transcript": context,
                    "reps": ", ".join(reps),
                    "clients": ", ".join(clients),
                })
                
                # Handle Gemini response
                text = response.content if hasattr(response, 'content') else str(response)
                    
            except Exception as e:
                text = f"[LLM invocation error: {e}]"
                print(f"[ERROR] LLM error while processing paragraph {idx+1}: {e}")

            # Try parsing JSON response
            parsed = None
            try:
                # Clean up response text - sometimes LLMs add markdown code blocks
                cleaned_text = text.strip()
                if cleaned_text.startswith("```json"):
                    cleaned_text = cleaned_text.replace("```json", "").replace("```", "").strip()
                elif cleaned_text.startswith("```"):
                    cleaned_text = cleaned_text.replace("```", "").strip()
                
                parsed = json.loads(cleaned_text)
                
                # Debug: Print what LLM decided
                is_faq = parsed.get("is_faq", "no")
                print(f"[DEBUG] Paragraph {idx+1}: is_faq={is_faq}, category={parsed.get('category', 'N/A')}")
                
            except Exception as parse_error:
                print(f"[WARN] Failed to parse JSON for paragraph {idx+1}: {parse_error}")
                print(f"[DEBUG] Raw LLM response: {text[:200]}...")
                parsed = None

            # Handle both boolean True and string "yes"
            is_faq_value = parsed.get("is_faq") if isinstance(parsed, dict) else None
            is_actually_faq = (
                is_faq_value == "yes" or 
                is_faq_value == True or 
                str(is_faq_value).lower() == "true"
            )
            
            # Only process and save if it's actually an FAQ
            if isinstance(parsed, dict) and is_actually_faq:
                # Build row with properly formatted question and responses
                row = {
                    "Original_Question": paragraph_text,  # Raw conversation
                    "Improved_Question": parsed.get("improved_question", paragraph_text),  # Formatted question
                    "Original_Response": parsed.get("original_response", ""),
                    "Improved_Response": parsed.get("improved_response", ""),
                    "Category": parsed.get("category", ""),
                    "Date": meeting_date,
                    # Keep some metadata for debugging
                    "Meeting_ID": meeting_id,
                    "Client": client_name,
                    "raw_analysis": text,
                }
                
                results.append(row)
                # Save incrementally
                self.writer.append_to_excel(row, "faq_knowledge.xlsx", sheet_name="FAQ")
                print(f"[SUCCESS] Saved FAQ #{len(results)}: {row['Category']}")
            else:
                if isinstance(parsed, dict):
                    print(f"[SKIP] Not an FAQ - Reason: {parsed.get('reason', 'No reason provided')}")
                else:
                    print(f"[SKIP] Failed to parse LLM response")
            
            # Progress indicator
            if (idx + 1) % 5 == 0:
                print(f"[INFO] Processed {idx + 1}/{len(paragraphs)} paragraphs")

        print(f"[INFO] Completed processing. Found {len(results)} FAQs for {client_name}")
        return results


if __name__ == "__main__":
    # Check Google API key
    google_api_key = os.getenv("GOOGLE_API_KEY")
    
    if not google_api_key:
        raise ValueError(
            "GOOGLE_API_KEY not found in environment variables.\n"
            "Please add it to your .env file:\n"
            "GOOGLE_API_KEY=your_api_key_here\n\n"
            "Get a free API key from: https://aistudio.google.com/app/apikey"
        )
    
    print(f"[INFO] Google API Key: ✓")
    print("[INFO] Initializing FAQ Agent with Google Gemini...")
    
    # Try different model names in order of preference
    models_to_try = [
        "gemini-2.5-flash",      # Latest
        "gemini-1.5-flash-latest",  # Stable latest
        "gemini-1.5-flash-001",  # Versioned stable
        "gemini-pro"              # Fallback
    ]
    
    agent = None
    for model_name in models_to_try:
        try:
            print(f"[INFO] Trying model: {model_name}...")
            agent = FAQAgent(model=model_name)
            break  # Success!
        except Exception as e:
            print(f"[WARN] Failed to initialize with {model_name}: {e}")
            continue
    
    if agent is None:
        raise RuntimeError(
            "Failed to initialize with any Gemini model. "
            "Please check your API key and internet connection."
        )
    
    # Example test run (replace with actual meeting_id and client_name)
    test_meeting_id = "M001"
    test_client_name = "Advik Hi Tech"
    faq_results = agent.process(test_meeting_id, test_client_name)
    print(f"\n[SUCCESS] Extracted {len(faq_results)} FAQ entries.")