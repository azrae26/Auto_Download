from datetime import datetime
from queue import Queue

debug_queue = Queue()

def debug_print(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f"[{timestamp}] {message}"
    debug_queue.put(formatted_message)
    print(formatted_message) 