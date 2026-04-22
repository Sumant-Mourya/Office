import sys
import os
import math
import subprocess
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QScrollArea,
    QGridLayout, QSizePolicy
)
from PySide6.QtCore import Qt


APP_FOLDER = "allapps"


class SmartUI(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("App Panel")
        self.resize(500, 400)

        self.main_layout = QVBoxLayout()

        self.title = QLabel("Applications")
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setStyleSheet("font-size:18px; font-weight:bold;")

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)

        self.container = QWidget()
        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.container.setLayout(self.grid)
        self.scroll.setWidget(self.container)

        self.main_layout.addWidget(self.title)
        self.main_layout.addWidget(self.scroll)

        self.setLayout(self.main_layout)

        self.center_window()
        self.load_buttons()

    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - 500) // 2
        y = (screen.height() - 400) // 2
        self.move(x, y)

    def load_buttons(self):
        base_path = os.path.join(os.getcwd(), APP_FOLDER)

        if not os.path.exists(base_path):
            os.makedirs(base_path)

        files = [f for f in os.listdir(base_path) if f.endswith(".py")]
        total = len(files)

        if total == 0:
            return

        cols = math.ceil(math.sqrt(total))
        rows = math.ceil(total / cols)

        cols = min(cols, 8)
        rows = min(rows, 10)

        # 🔥 SMALL BUTTON MODE
        small_mode = total <= 3

        index = 0

        for r in range(rows):
            items_in_row = min(cols, total - index)
            start_col = (cols - items_in_row) // 2

            for c in range(items_in_row):
                file = files[index]
                file_path = os.path.join(base_path, file)

                btn = QPushButton(file.replace(".py", ""))

                if small_mode:
                    btn.setFixedSize(150, 50)   # 👈 small buttons
                else:
                    btn.setMinimumHeight(70)
                    btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

                # ✅ FIXED CLICK (no lambda bug)
                btn.clicked.connect(self.make_runner(file_path))

                self.grid.addWidget(btn, r, start_col + c)

                index += 1

    # ✅ Proper function binding
    def make_runner(self, path):
        def run():
            subprocess.Popen([sys.executable, path])
        return run


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SmartUI()
    window.show()
    sys.exit(app.exec())