import sys
import os
import subprocess
import socket
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout,
    QLabel, QPushButton
)
from PySide6.QtCore import Qt, QThread, Signal


TARGET_SCRIPT = "Allappui.py"


def is_internet_available():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except:
        return False


class UpdateWorker(QThread):
    finished = Signal(bool)

    def run(self):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    
            commands = [
                ["git", "fetch", "--all"],
                ["git", "reset", "--hard", "origin/main"],
                ["git", "clean", "-fd"]
            ]
    
            for cmd in commands:
                process = subprocess.Popen(
                    cmd,
                    cwd=os.getcwd(),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
    
                try:
                    process.communicate(timeout=15)  # ⏱️ prevent hang
                except subprocess.TimeoutExpired:
                    process.kill()
                    self.finished.emit(False)
                    return
    
            self.finished.emit(True)
    
        except Exception as e:
            print("Error:", e)
            self.finished.emit(False)


class UpdateDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Updater")
        self.setFixedSize(360, 200)

        # 🔥 Improved modern UI
        self.setStyleSheet("""
            QDialog {
                background-color: #121212;
            }

            QLabel {
                color: #ffffff;
                font-size: 18px;
                font-weight: 600;
            }

            QPushButton {
                background-color: #3a86ff;
                color: white;
                font-size: 15px;
                font-weight: bold;
                border-radius: 10px;
                padding: 10px;
                min-height: 40px;
            }

            QPushButton:hover {
                background-color: #265df2;
            }

            QPushButton:pressed {
                background-color: #1f4ed8;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(20)

        self.label = QLabel("Checking for updates...")
        self.label.setAlignment(Qt.AlignCenter)

        self.button = QPushButton("Retry")
        self.button.setVisible(False)
        self.button.clicked.connect(self.retry)

        layout.addStretch()
        layout.addWidget(self.label)
        layout.addWidget(self.button)
        layout.addStretch()

        self.setLayout(layout)

        self.center_window()
        self.start_update()

    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2
        )

    def start_update(self):
        if not is_internet_available():
            self.label.setText("No Internet Connection")
            self.button.setVisible(True)
            return

        self.label.setText("Updating...")
        self.button.setVisible(False)

        self.worker = UpdateWorker()
        self.worker.finished.connect(self.update_done)
        self.worker.start()

    def retry(self):
        self.label.setText("Retrying...")
        self.button.setVisible(False)
        self.start_update()

    def update_done(self, success):
        if not success:
            self.label.setText("Update Failed")
            self.button.setVisible(True)
            return

        self.label.setText("Launching...")

        script_path = os.path.join(os.getcwd(), TARGET_SCRIPT)

        if os.path.exists(script_path):
            subprocess.Popen([sys.executable, script_path])

        QApplication.quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dialog = UpdateDialog()
    dialog.show()
    sys.exit(app.exec())