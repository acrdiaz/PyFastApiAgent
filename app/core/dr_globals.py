import os
from pathlib import Path
from pydantic import SecretStr


# --- Constants ---
DR_API = "Hi, master"

# --- Service Config ---
DR_STATUS_LOG_INTERVAL = 10  # seconds
DR_QUEUE_TIMEOUT = 1.0  # seconds
DR_THREAD_JOIN_TIMEOUT = 5.0  # seconds
DR_POLLING_INTERVAL = 0.95  # seconds prevent CPU hogging

# --- Path Config ---
DR_BASE_PATH = Path(__file__).parent.parent.absolute()
DR_DB_FOLDER = "db"
DR_PROMPT_FILE = "dr_prompt.txt"
DR_RESPONSE_FILE = "dr_response.txt"

DR_PROMPT_FILE_PATH = DR_BASE_PATH / DR_DB_FOLDER / DR_PROMPT_FILE
DR_RESPONSE_FILE_PATH = DR_BASE_PATH / DR_DB_FOLDER / DR_RESPONSE_FILE

# --- Browser Config ---
DR_BROWSER = "chrome"
MAX_STEPS = 30
MAX_ACTIONS_PER_STEP = 10

# --- LLM Config ---
LLM_GEMINI = 'gemini-2.0-flash'
LLM_GPT = 'gpt-4o-mini'
LLM_DEFAULT = LLM_GEMINI
API_KEY = SecretStr(os.getenv("GEMINI_API_KEY")) # type: ignore

# --- Secrets ---

# --- Agent status ---
class DRGlobals:
    def __init__(self):
        self.DR_AGENT_RUNNNG = False

# Singleton for reuse (avoids recreating class)
# if DR_GLOBALS:
#     print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ DR_globals.py")
#     DR_GLOBALS = DRGlobals()

print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ DR_globals.py")
DR_GLOBALS = DRGlobals()