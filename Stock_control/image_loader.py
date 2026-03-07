from PyQt5.QtCore import QRunnable, pyqtSignal, QObject
import os

class WorkerSignals(QObject):
    finished = pyqtSignal(object)

class ImageLoaderRunnable(QRunnable):
    def __init__(self, folder):
        super().__init__()
        self.folder = folder
        self.signals = WorkerSignals()

    def run(self):
        try:
            image_files = {
                os.path.splitext(f[:10])[0]: os.path.join(self.folder, f)
                for f in os.listdir(self.folder)
                if f.lower().endswith((".png", ".jpg"))
            }
        except Exception as e:
            image_files = {}
        self.signals.finished.emit(image_files)