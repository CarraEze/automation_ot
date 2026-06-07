from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler
import subprocess
import time

# Folder to monitor
WATCH_FOLDER = r"\\10.2.2.24\ecarrazana"

# Script to execute when a change happens
SCRIPT_TO_RUN = r"C:\Users\ecarrazana\Desktop\python\importSheetsV5.exe"


# Store last execution time per file
last_execution = {}

# Ignore repeated events within X seconds
COOLDOWN = 5


class FolderHandler(FileSystemEventHandler):

    def process(self, event):

        if event.is_directory:
            return

        file_path = event.src_path

        now = time.time()

        # Check last execution time
        if file_path in last_execution:

            elapsed = now - last_execution[file_path]

            if elapsed < COOLDOWN:
                return

        # Update timestamp
        last_execution[file_path] = now

        print(f"Processing: {file_path}")

        subprocess.run([SCRIPT_TO_RUN, file_path])

    def on_created(self, event):
        self.process(event)

    def on_modified(self, event):
        self.process(event)


observer = PollingObserver(timeout=2)

handler = FolderHandler()

observer.schedule(handler, WATCH_FOLDER, recursive=False)

observer.start()

print("Watching network folder...")

try:
    while True:
        time.sleep(1)

except KeyboardInterrupt:
    observer.stop()

observer.join()