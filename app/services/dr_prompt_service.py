# from app.services.dr_browser_agent import DRBrowserUseAgent
from app.services.dr_prompt_consumer import DRPromptConsumer
from app.services.dr_prompt_producer import DRPromptProducer
from app.core.dr_task_queue import DRTaskQueue
print("##################################################################################### dr_prompt_service.py")
from app.core.dr_globals import (
    DR_PROMPT_FILE_PATH,
    DR_BROWSER,
)
print("##################################################################################### dr_prompt_service.py")
from app.services.dr_web_agent_worker import DRWebAgentWorker
from app.utils.dr_utils_file import DRUtilsFile
from typing import Dict, Any, Optional, List, Callable

import asyncio
import logging
import threading
import time


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('PromptService')


class DRPromptService:
    """
    A service that runs in the background and processes a queue of prompts.
    """
    def __init__(self, processor_func: Callable[[str, Dict[str, Any]], Any] = None): # type: ignore
        """
        Initialize the prompt service.
        
        Args:
            processor_func: Function to process prompts. Should accept (prompt_text, metadata)
                           If None, a default processor that just logs the prompt will be used.
        """
        self.running = False
        self.results = {}  # Store results by prompt_id
        self.processor = processor_func or self._default_processor

        self.prompt_file = DRUtilsFile(DR_PROMPT_FILE_PATH)

        self.prompt_queue = DRTaskQueue()  # Single queue for all prompt processing
        self.web_agent_worker = DRWebAgentWorker(DR_BROWSER)
        self.producer = DRPromptProducer(self.prompt_queue, self.prompt_file)
        self.consumer = DRPromptConsumer(self.prompt_queue, self.web_agent_worker)

        logger.info("DRPromptService initialized")

    def _default_processor(self, prompt: str, metadata: Dict[str, Any]) -> str:
        """Default prompt processor that just logs the prompt."""
        logger.info(f"Processing prompt: {prompt[:50]}...")
        logger.info(f"Metadata: {metadata}")

        # Using DRBrowserUseAgent from the separate file
        # browserAgent = DRBrowserUseAgent(browser_name="edge") # AA1 global
        # asyncio.run(browserAgent.main(prompt, metadata))

        # Use metadata in the response if available
        priority = metadata.get('priority', 'normal')
        source = metadata.get('source', 'unknown')

        return f"Processed: {prompt[:50]}... (Priority: {priority}, Source: {source})"

    async def _run_browser_agent(self, agent):
        """Run the browser agent and close the browser when done."""
        try:
            await agent.run()
        finally:
            await agent.browser.close()

    def _process_prompt_and_update_result(self, prompt_id, prompt_text, metadata):
        """Process a prompt and update its result."""
        try:
            # Process the prompt
            result = self.processor(prompt_text, metadata)
            status = "completed"
        except Exception as e:
            logger.error(f"Error processing prompt {prompt_id}: {str(e)}")
            result = str(e)
            status = "error"
        
        # Store the result
        self.results[prompt_id] = {
            "status": status,
            "result": result,
            "completed_at": time.time(),
            "prompt": prompt_text,
            "metadata": metadata
        }

    def start(self):
        """Start the service."""
        if self.running:
            logger.warning("Service is already running")
            return
        
        self.running = True
        self.producer.start()
        self.consumer.start()
        logger.info("Service started")

    def stop(self):
        """Stop the service."""
        if not self.running:
            logger.warning("Service is not running")
            return
        
        logger.info("Stopping service...")
        self.running = False
        self.producer.stop()
        self.consumer.stop()
        # self.producer.join() # AA1 timeout=DR_THREAD_JOIN_TIMEOUT
        # self.consumer.join()
        logger.info("Service stopped")

    def get_status(self, prompt_id: str) -> Optional[Dict[str, Any]]:
        """
        Get the status of a prompt.
        
        Args:
            prompt_id: The ID of the prompt
            
        Returns:
            A dictionary with the status information, or None if not found
        """
        return self.results.get(prompt_id)

    def get_all_statuses(self) -> Dict[str, Dict[str, Any]]:
        """Get the status of all prompts."""
        return self.results