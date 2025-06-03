print("#################################################################################### consumer.py")
from app.core.dr_globals import DR_POLLING_INTERVAL
print("#################################################################################### consumer.py")
from app.core.dr_task_queue import DRTaskQueue
from app.services.dr_web_agent_worker import DRWebAgentWorker
from typing import Optional

import asyncio
import threading
import logging
import time


logger = logging.getLogger('PromptService')

class DRPromptConsumer(threading.Thread):
    def __init__(self, queue, worker: Optional[DRWebAgentWorker] = None):
        super().__init__(daemon=True)
        self.queue = queue
        self.worker = worker or DRWebAgentWorker()
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()
        self.worker.shutdown()

    def _handle_result(self, future):
        """Callback for completed tasks"""
        try:
            result = future.result()
            logging.info(f"Consumer received result: {result}")
        except Exception as e:
            logging.error(f"Task failed: {e}")

    def run(self):
        while not self._stop_event.is_set():
            try:
                # Small delay to prevent CPU hogging
                time.sleep(DR_POLLING_INTERVAL)

                if self.queue.empty():
                    continue

                prompt_id, prompt_text, metadata = self.queue.get()

                try:
                    # Process the prompt
                    # await self.worker.process(prompt_text, self._handle_result)
                    # browserAgent = DRWebAgentWorker()
                    asyncio.run(self.worker.main(prompt_text))

                except Exception as e:
                    logger.error(f"Error processing prompt {prompt_id}: {e}")

                self.queue.task_done()
                
            except Exception as e:
                logger.error(f"Error in consumer thread: {str(e)}")