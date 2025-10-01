import tiktoken
from typing import List, Dict, Any


class TranscriptChunker:
    """
    Utility to split transcripts into token-safe chunks and support hierarchical summarization.
    """

    def __init__(self, model_name: str = "gpt-3.5-turbo"):
        self.encoding = tiktoken.encoding_for_model(model_name)
        self.max_chunk_tokens = 1800  # safe size per chunk

    def count_tokens(self, text: str) -> int:
        """Count tokens in a string using model encoding."""
        return len(self.encoding.encode(text))

    def chunk_transcript(self, transcript: List[Dict[str, Any]]) -> List[str]:
        """
        Split transcript into chunks of ~1800 tokens.
        transcript: list of {speaker, start, sentence}
        Returns list of transcript chunks as plain text.
        """
        chunks, current_chunk, current_tokens = [], [], 0

        for entry in transcript:
            speaker = entry.get("speaker", "Unknown")
            time = entry.get("start", "")
            sentence = entry.get("sentence", "")
            line = f"{speaker} [{time}]: {sentence}\n"

            line_tokens = self.count_tokens(line)

            if current_tokens + line_tokens > self.max_chunk_tokens:
                chunks.append("".join(current_chunk))
                current_chunk, current_tokens = [], 0

            current_chunk.append(line)
            current_tokens += line_tokens

        if current_chunk:
            chunks.append("".join(current_chunk))

        return chunks

    def filter_by_speaker(self, transcript: List[Dict[str, Any]], include_keywords: List[str]) -> List[Dict[str, Any]]:
        """
        Filter transcript by speakers (e.g., Dice reps).
        include_keywords: list of substrings to match in speaker names.
        """
        filtered = []
        for entry in transcript:
            speaker = entry.get("speaker", "")
            if any(k.lower() in speaker.lower() for k in include_keywords):
                filtered.append(entry)
        return filtered

    def summarize_chunks(self, chunks: List[str], llm_chain, extra_inputs: Dict[str, str] = None) -> str:
        """
        Summarize long transcript hierarchically.
        - Summarize each chunk individually with LLM.
        - Merge summaries into final summary.
        llm_chain: LangChain LLMChain for summarization.
        extra_inputs: dict of extra keys like {"reps": "...", "clients": "..."}.
        """
        sub_summaries = []
        for i, chunk in enumerate(chunks, 1):
            try:
                inputs = {"transcript": chunk}
                if extra_inputs:
                    inputs.update(extra_inputs)
                summary = llm_chain.invoke(inputs)
                text = summary["text"] if isinstance(summary, dict) else summary
                sub_summaries.append(f"Chunk {i}: {text}")
            except Exception as e:
                sub_summaries.append(f"Chunk {i}: [Error summarizing chunk: {e}]")

        # Merge summaries into final
        merged_input = "\n".join(sub_summaries)
        final_inputs = {"transcript": merged_input}
        if extra_inputs:
            final_inputs.update(extra_inputs)

        final_summary = llm_chain.invoke(final_inputs)
        return final_summary["text"] if isinstance(final_summary, dict) else final_summary
