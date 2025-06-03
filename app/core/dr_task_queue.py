import queue

class DRTaskQueue:
    def __init__(self):
        self.queue = queue.Queue()
    
    def put(self, item):
        self.queue.put(item)
    
    def get(self):
        return self.queue.get()
    
    def task_done(self):
        self.queue.task_done()
    
    def empty(self):
        return self.queue.empty()
    
    def get_queue_size(self):
        return self.queue.qsize()