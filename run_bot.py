from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import time

BOT_FILE = "bot.py"

class ReloadHandler(FileSystemEventHandler):
    def __init__(self):
        self.proc = None
        self.start_bot()

    def start_bot(self):
        if self.proc:
            self.proc.terminate()
        print("Starting bot...")
        self.proc = subprocess.Popen(["python", BOT_FILE])

    def on_modified(self, event):
        if event.src_path.endswith(".py"):
            print(f"Changed: {event.src_path} -> restarting bot")
            self.start_bot()

if __name__ == "__main__":
    handler = ReloadHandler()
    observer = Observer()
    observer.schedule(handler, path=".", recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    finally:
        observer.stop()
        observer.join()
        if handler.proc:
            handler.proc.terminate()
