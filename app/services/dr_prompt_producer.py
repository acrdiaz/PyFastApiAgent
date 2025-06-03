print("##################################################################################### producer.py")
from app.core.dr_globals import DR_POLLING_INTERVAL
print("##################################################################################### producer.py")

import threading
import logging
import time


logger = logging.getLogger('PromptService')

class DRPromptProducer(threading.Thread):
    def __init__(self, queue, prompt_file):
        super().__init__(daemon=True)
        self.queue = queue
        self.promptFile = prompt_file
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def run(self):
        while not self._stop_event.is_set():
            try:
                # Small delay to prevent CPU hogging
                time.sleep(DR_POLLING_INTERVAL)

                if self.promptFile.get_file_size() == 0:
                    continue

                prompt_text = self.promptFile.load_file()
                if not prompt_text:
                    continue

                logger.info(f"Loaded prompt: {prompt_text}")
                self.promptFile.clean_file()
                
                metadata = {'priority': 'Normal'}
                prompt_id = f"prompt_{int(time.time() * 1000)}"

                self.queue.put((prompt_id, prompt_text, metadata))

            except Exception as e:
                logger.error(f"Error in producer thread: {str(e)}")