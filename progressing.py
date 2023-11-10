import threading
import time

# Fortschrittsanzeige mit drei Punkten nebeneinander
class ProgressIndicator:
    def __init__(self):
        self.stopped = False

    def start(self):
        self.thread = threading.Thread(target=self.run)
        self.thread.start()

    def stop(self):
        self.stopped = True
        self.thread.join()

    def run(self):
        while not self.stopped:
            for i in range(3):
                print(".", end="", flush=True)
                time.sleep(0.5)
            print("\b\b\b   \b\b\b", end="", flush=True) # Löschen der Punkte
            time.sleep(0.5)
