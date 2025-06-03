from typing import Optional
from functools import lru_cache

class DRSystemPrompts:
    """Zero-I/O, high-performance prompt manager."""

    # --- Simple Prompts (Constants) ---
    DIRECT_QUESTION = "Prompt: {prompt}."
    SUMMARIZE = "Summarize this in {sentences_number}. Sentences: {text}"

    # --- Dynamic Prompts (Methods) ---
    @staticmethod
    def translate(text: str, language: str = "English") -> str:
        return f"Translate '{text}' to {language}."

    @staticmethod
    def direct_llm_question(text: str) -> str:
        return (
            f"Prompt: {text}.\n"
            "Communicate the answer properly, based on the question.\n"
            "If the Prompt is a simple knowledge query, answer it.\n"
            "If the Prompt looks like instructions to interact with a website, respond with 'NEEDS_BROWSER'."
        )
    
    @staticmethod
    def prompt_additional_information(text: str) -> str:
        return (
            ""
        )

    @staticmethod
    def prompt_error_recovery() -> str:
        return (
            "If a click fails, proceed:\n"
            "1. Verify is enabled and remember its text.\n"
            "2. If enabled retry once.\n"
            "3. If still unresponsive:\n"
            "   - Report: 'Help needed: Clicked ELEMENT '[text]', but no action.'\n"
            "   - Finish Agent.\n"
        )

    # --- Heavy Prompts (Cached) ---
    @classmethod
    @lru_cache(maxsize=10)  # Cache if repeatedly used
    def few_shot_example(cls, style: str = "formal") -> str:
        # Large prompt (>1KB) stored in-code (no I/O)
        templates = {
            "formal": "You are an expert...",
            "casual": "Hey AI, can you..."
        }
        return templates[style]

# Singleton for reuse (avoids recreating class)
print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ DR_system_prompts.py")
DR_SYSTEM_PROMPTS = DRSystemPrompts()