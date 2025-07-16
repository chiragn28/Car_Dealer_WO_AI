import sys
import os
import threading
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QPushButton, QTextEdit,
    QVBoxLayout, QWidget, QFileDialog
)
from PyQt6.QtCore import Qt

from utils import ConfigManager, AudioProcessor, DataExtractor, ExcelManager, EmailSender

# Load config and components
CONFIG_PATH = "c:/Users/Administrator/Desktop/Chirag/car_dealer/config.ini"
config = ConfigManager(CONFIG_PATH)
audio_processor = AudioProcessor(config.get_whisper_config("model_path"))
data_extractor = DataExtractor(config)
excel_manager = ExcelManager(config.get_path("appointments_file"))
email_sender = EmailSender(config)

class DragDropLabel(QLabel):
    def __init__(self, parent):
        super().__init__("📂 Drag and Drop Audio File Here", parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("QLabel { border: 2px dashed #aaa; font-size: 16px; padding: 40px; }")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith(('.mp3', '.wav', '.mp4')):
                self.parent().process_file(file_path)
                break


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("🎧 Appointment Extractor")
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout()
        self.drag_drop_label = DragDropLabel(self)
        self.drag_drop_label.setAcceptDrops(True)

        self.browse_button = QPushButton("📁 Browse Audio File")
        self.browse_button.clicked.connect(self.browse_file)

        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        self.output_box.setStyleSheet("font-family: Courier;")

        layout.addWidget(self.drag_drop_label)
        layout.addWidget(self.browse_button)
        layout.addWidget(self.output_box)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Audio File", "", "Audio Files (*.mp3 *.wav *.mp4)")
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        self.output_box.clear()
        self.output_box.append(f"🔄 Processing: {file_path}")
        thread = threading.Thread(target=self.run_pipeline, args=(file_path,), daemon=True)
        thread.start()

    def run_pipeline(self, file_path):
        try:
            transcript = audio_processor.transcribe_audio(file_path)
            self.append_output("✅ Transcription complete\n")
            self.append_output(transcript + "\n\n")

            data = data_extractor.extract_info(transcript)
            self.append_output("✅ Field extraction complete\n")

            for k, v in data.items():
                self.append_output(f"{k}: {v}")

            excel_manager.add_appointment(data)
            self.append_output("\n✅ Saved to Excel")

            recipient = data.get("email") or "cnayak70@gmail.com"
            email_sender.send_confirmation(recipient, data)
            self.append_output(f"\n✅ Email sent to {recipient}")

            processed_path = os.path.join(config.get_path("processed_dir"), os.path.basename(file_path))
            os.rename(file_path, processed_path)
            self.append_output(f"\n✅ Moved to {processed_path}")
        except Exception as e:
            self.append_output(f"\n❌ Error: {str(e)}")

    def append_output(self, text):
        self.output_box.append(text)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
