import asyncio
import logging
import os
import time
# from typing import Optional

print("####################################################################################A DR_web_agent_worker.py")
from app.core.dr_globals import (
    API_KEY,
    LLM_DEFAULT,
    DR_RESPONSE_FILE_PATH,
    MAX_STEPS,
    LLM_GEMINI,
    LLM_GPT,
    LLM_DEFAULT,
    DR_GLOBALS,
)
print("####################################################################################B DR_web_agent_worker.py")
from app.core.dr_system_prompts import DR_SYSTEM_PROMPTS, DR_SYSTEM_PROMPTS
from app.utils.dr_utils_file import DRUtilsFile

from browser_use import Agent, Browser, BrowserConfig
from concurrent.futures import Executor, ThreadPoolExecutor
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_openai import ChatOpenAI


class DRWebAgentWorker:
    def __init__(self,
                 browser_name: str = "chrome",
                 headless: bool = False,
                 executor: Executor = None): # type: ignore
        self._loop = asyncio.new_event_loop()
        self.executor = executor or ThreadPoolExecutor(max_workers=4)
        self._running = False

        self._response_file = DRUtilsFile(DR_RESPONSE_FILE_PATH)

        if browser_name.lower() == "chrome":
            path_32 = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            path_64 = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        elif browser_name.lower() == "edge":
            path_32 = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            path_64 = "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe"
        else:
            raise FileNotFoundError(f"Browser not supported: {browser_name}")

        browser_path = self._app_exits(path_32, path_64)

        if browser_path is None:
            raise FileNotFoundError(f"Browser path file not found: {browser_name}")

        self.browser = Browser(
            config=BrowserConfig(
                chrome_instance_path=browser_path,
                headless=headless,
            )
        )

    def _app_exits(self, path_32, path_64):
        """Check if the application exists in the given paths.

        Args:
            folder_path_32 (str): Path to the 32-bit application
            folder_path_64 (str): Path to the 64-bit application

        Returns:
            str: Path of the existing application or None
        """
        if os.path.exists(path_32):
            return path_32
        elif os.path.exists(path_64):
            return path_64
        else:
            raise FileNotFoundError(f"Browser executable not found.")

    def _create_llm(self, llm_model: str = LLM_DEFAULT):
        if not API_KEY:
            # logging.info("No API key provided for LLM")
            raise ValueError("No API key provided for LLM")

        if llm_model == LLM_GEMINI:
            return ChatGoogleGenerativeAI(
                model=llm_model,
                api_key=API_KEY,
            )
        elif llm_model == LLM_GPT:
            return ChatOpenAI(
                model=llm_model,
            )
        else:
            raise ValueError(f"Unsupported LLM model: {llm_model}")

    def _save_browser_response(self, page_content):
        last_result_content = page_content.final_result()
        self._save_text_response(last_result_content)

    def _save_text_response(self, text):
        DR_GLOBALS.DR_AGENT_RUNNNG = False
        self._response_file.write_file(text)
        logging.info(f"Response saved to {self._response_file.file_path}")

    async def direct_llm_question(self, prompt: str, llm):
        """Get direct response from LLM for simple knowledge queries."""
        try:
            additional_information = DR_SYSTEM_PROMPTS.prompt_additional_information("Direct Question")

            prompt_combined = f"{prompt}.\n{additional_information or ''}"

            prompt_llm = DR_SYSTEM_PROMPTS.direct_llm_question(prompt_combined)
            # logging.info(f"Prompt for LLM: {prompt_llm}")
            response = llm.invoke(prompt_llm)
            answer = response.content.strip() # type: ignore

            if answer != "NEEDS_BROWSER":
                self._save_text_response(answer)

            return None if "NEEDS_BROWSER" in answer else answer

        except Exception as e:
            print(f"Error getting LLM response: {e}")
            return None

    async def browse_web(self, prompt: str, llm):
        additional_information = DR_SYSTEM_PROMPTS.prompt_additional_information("")

        # Add context about error recovery
        error_recovery_context = DR_SYSTEM_PROMPTS.prompt_error_recovery()
        
        prompt_combined = f"{prompt}.\n{additional_information or ''}\n{error_recovery_context or ''}"

        # print(f"🧠 Combined prompt: {prompt_combined[:50]}...")  # Debugging output

        agent = Agent(
            task=prompt_combined,
            # message_context=combined_context,
            # planner_llm=planner_llm,
            # planner_interval=1,
            llm=llm,
            browser=self.browser,
            # max_actions_per_step=MAX_ACTIONS_PER_STEP,
            max_failures=1,
            retry_delay=1,                            # Short delay in seconds between retries
        )
        
        time_start = time.time()
        logging.info(f"Running browser agent with prompt: {prompt[:50]}...")

        page_content = await agent.run(max_steps=MAX_STEPS)
        self._save_browser_response(page_content)

        time_duration = time.time() - time_start
        logging.info(f"Agent completed in {time_duration:.2f} seconds")

        await agent.browser.close() # type: ignore
        await self.browser.close() # type: ignore # AA1 is this needed?

    async def main(self, prompt: str, llm_model: str = LLM_DEFAULT):
        DR_GLOBALS.DR_AGENT_RUNNNG = True
        print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
        print(f"DR_AGENT_RUNNNG: {DR_GLOBALS.DR_AGENT_RUNNNG}")

        logging.info(f"🧠 Agent found a prompt: {prompt[:50]}...")

        llm = self._create_llm(llm_model)

        # Agent A
        logging.info(f"🧠 Agent A - direct question")
        answer = await self.direct_llm_question(prompt, llm)
        if answer != None:
            return

        # Agent B
        logging.info(f"🧠 Agent B - browse the web")
        await self.browse_web(prompt, llm)

    # def process(self, prompt, callback=None):
    #     # """Submit prompt for processing"""
    #     # future = self.executor.submit(self._execute_agent_work, prompt, self.browser)
    #     # if callback:
    #     #     future.add_done_callback(callback)
    #     # return future
    #     """Bridge sync world to async"""
    #     async def _run_and_notify():
    #         result = await self._execute_agent_work(prompt)
    #         if callback:
    #             callback(result)
    #         return result

    #     future = asyncio.run_coroutine_threadsafe(
    #         _run_and_notify(),
    #         self._loop
    #     )
    #     return future

    # @staticmethod
    # async def _execute_agent_work(prompt="", browser=None):
    #     if browser is None: 
    #         raise ValueError("Browser instance is required")
    #     if not prompt:
    #         raise ValueError("Prompt is required")
        
    #     logging.info(f"Agent started processing: {prompt}")

    #     llm = ChatGoogleGenerativeAI(
    #         model=DEFAULT_MODEL,
    #         api_key=API_KEY,
    #     )

    #     agent = Agent(
    #         task=prompt,
    #         llm=llm,
    #         browser=browser,
    #     )
        
    #     await agent.run()
    #     await agent.browser.close()

    #     return f"Processed: {prompt}"

    # def start(self):
    #     """Start event loop in background thread"""
    #     self._running = True
    #     threading.Thread(
    #         target=self._run_event_loop,
    #         daemon=True
    #     ).start()

    # def _run_event_loop(self):
    #     """Run the event loop in dedicated thread"""
    #     asyncio.set_event_loop(self._loop)
    #     while self._running:
    #         self._loop.run_forever()

    def shutdown(self):
        """Cleanup resources"""
        self._running = False
        self._loop.call_soon_threadsafe(self._loop.stop)
        self.executor.shutdown(wait=True)

if __name__ == "__main__":
    prompt = "open https://espanol.yahoo.com/"
    buService = DRWebAgentWorker("chrome")
    asyncio.run(buService.main(prompt))