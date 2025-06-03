print("#################################################################################### main.py")
from app.core.dr_globals import (
    DR_API,
    DR_PROMPT_FILE_PATH,
    DR_RESPONSE_FILE_PATH,
    DR_GLOBALS
)
print("#################################################################################### main.py")
from app.utils.dr_utils_file import DRUtilsFile

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from pydantic import BaseModel

# Import routers when you create them
# from app.api.example_router import router as example_router


app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

PROMPT_FILE = DRUtilsFile(DR_PROMPT_FILE_PATH)
RESPONSE_FILE = DRUtilsFile(DR_RESPONSE_FILE_PATH)

# Register routers
# app.include_router(example_router)

class PromptRequest(BaseModel):
    prompt: str

@app.get("/")
async def root():
    return {"message": f"Welcome to the FastAPI application: {DR_API}"}

@app.post("/prompt/")
async def create_prompt(request: PromptRequest):
    prompt = request.prompt.strip()
    if not prompt:
        return {"message": "Prompt cannot be empty."}
    if PROMPT_FILE.get_file_size() == 0:
        PROMPT_FILE.write_file(prompt)
        return {"message": "Prompt created successfully!", "prompt": prompt}
    else:
        return {"message": "Try again later, a prompt is in queue."}

@app.post("/clearPromptResponse/")
async def clear_prompt_response():
    if PROMPT_FILE.get_file_size() > 0:
        PROMPT_FILE.clean_file()

    if RESPONSE_FILE.get_file_size() > 0:
        RESPONSE_FILE.clean_file()

    return {"message": "Prompt, response cleared successfully!"}

@app.get("/response/")
async def get_response():
    print(f"DR_AGENT_RUNNNG: {DR_GLOBALS.DR_AGENT_RUNNNG}")

    if DR_GLOBALS.DR_AGENT_RUNNNG:
        return {"message": f"Please wait."}
        # return {"message": "Please wait, the agent is still processing."}

    if RESPONSE_FILE.get_file_size() > 0:
        text = RESPONSE_FILE.load_file()
        return {"message": f"{text}"}

    return {"message": f"No response available."}