import sys
import subprocess
import os
import time
import json
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QMessageBox, QProgressBar, QListWidget,
    QListWidgetItem, QSizePolicy
)
from PyQt5.QtCore import Qt, QTimer
import win32com.client
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

LOG_DIR = os.path.join(os.getenv("LOCALAPPDATA"), "SlideHelm", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

class OverlayWindow(QWidget):
    def __init__(self, on_close_callback):
        super().__init__()
        self.on_close_callback = on_close_callback

        self.setWindowTitle("SlideHelm Overlay")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setStyleSheet("background-color: rgba(30, 30, 30, 220); color: white; font-size: 14px;")
        self.setWindowOpacity(0.8)
        self.offset = None

        self.timer_label = QLabel("⏱️ Time left: 60:00")
        self.slide_label = QLabel("Slide 0 / 0")
        self.message_label = QLabel("✅ You're on track!")
        self.message_label.setWordWrap(True)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)

        row = QHBoxLayout()
        row.addWidget(self.timer_label)
        row.addWidget(self.slide_label)
        row.addWidget(self.message_label, stretch=1)
        row.addWidget(self.close_btn)

        layout = QVBoxLayout()
        layout.addLayout(row)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)
        self.setFixedSize(650, 80)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.offset = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if self.offset and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.offset)

    def mouseReleaseEvent(self, event):
        self.offset = None

    def closeEvent(self, event):
        self.on_close_callback()
        event.accept()

class LogViewer(QWidget):
    def __init__(self, log_dir):
        super().__init__()
        self.setWindowTitle("SlideHelm – Log Viewer")
        self.setMinimumSize(800, 600)
        self.setStyleSheet("background-color: #111; color: white; font-size: 12px;")
        self.log_dir = log_dir

        self.layout = QVBoxLayout()
        self.list_widget = QListWidget()
        self.layout.addWidget(QLabel("Presentation Logs:"))
        self.layout.addWidget(self.list_widget)
        self.setLayout(self.layout)

        self.refresh_logs()

    def refresh_logs(self):
        self.list_widget.clear()
        if not os.path.exists(self.log_dir):
            return

        files = sorted(os.listdir(self.log_dir), reverse=True)
        for file in files:
            if file.endswith(".json"):
                path = os.path.join(self.log_dir, file)
                try:
                    with open(path, "r") as f:
                        data = json.load(f)
                    item = QListWidgetItem()
                    summary = f"{file} | ⏱ {data.get('duration_minutes')} min | 📊 {data.get('slide_count')} slides"
                    widget = QWidget()
                    row = QHBoxLayout()
                    label = QLabel(summary)
                    view_btn = QPushButton("View")
                    delete_btn = QPushButton("Delete")

                    view_btn.clicked.connect(lambda _, f=path: self.view_log(f))
                    delete_btn.clicked.connect(lambda _, f=path: self.delete_log(f))

                    row.addWidget(label)
                    row.addWidget(view_btn)
                    row.addWidget(delete_btn)
                    widget.setLayout(row)
                    item.setSizeHint(widget.sizeHint())
                    self.list_widget.addItem(item)
                    self.list_widget.setItemWidget(item, widget)
                except Exception as e:
                    print(f"Error loading log {file}: {e}")

    def delete_log(self, filepath):
        try:
            os.remove(filepath)
            self.refresh_logs()
        except Exception as e:
            QMessageBox.critical(self, "Delete Failed", str(e))

    def view_log(self, filepath):
        with open(filepath, "r") as f:
            data = json.load(f)

        entries = data.get("entries", [])
        times = [entry["elapsed_seconds"] for entry in entries]
        slides = [entry["slide"] for entry in entries]

        slide_times = {}
        for i in range(1, len(entries)):
            slide = entries[i - 1]["slide"]
            delta = entries[i]["elapsed_seconds"] - entries[i - 1]["elapsed_seconds"]
            slide_times[slide] = slide_times.get(slide, 0) + delta

        avg_time = sum(slide_times.values()) / len(slide_times) if slide_times else 0
        longest_slide = max(slide_times, key=slide_times.get) if slide_times else None

        fig, ax = plt.subplots()
        ax.plot(list(slide_times.keys()), list(slide_times.values()), marker='o')
        ax.set_title("Slide Time Breakdown")
        ax.set_xlabel("Slide Number")
        ax.set_ylabel("Seconds Spent")

        stats = f"\nAverage time per slide: {avg_time:.2f}s"
        if longest_slide:
            stats += f"\nLongest on slide: {longest_slide} ({slide_times[longest_slide]}s)"

        QMessageBox.information(self, "Summary Stats", stats)
        fig.show()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SlideHelm")
        self.setStyleSheet("background-color: #222; color: white; font-size: 14px;")
        self.setFixedSize(400, 220)

        self.ppt = None
        self.presentation = None
        self.start_time = None
        self.presentation_minutes = 60
        self.slide_log = []

        self.overlay = OverlayWindow(self.end_session)
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update_overlay)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Set presentation duration (minutes):"))

        self.duration_input = QLineEdit()
        self.duration_input.setPlaceholderText("e.g. 60")
        layout.addWidget(self.duration_input)

        row = QHBoxLayout()

        launch_btn = QPushButton("Launch PowerPoint")
        launch_btn.clicked.connect(self.launch_powerpoint)
        row.addWidget(launch_btn)

        start_btn = QPushButton("Start Timer")
        start_btn.clicked.connect(self.start_timer)
        row.addWidget(start_btn)

        layout.addLayout(row)

        log_btn = QPushButton("View Logs")
        log_btn.clicked.connect(self.show_log_viewer)
        layout.addWidget(log_btn)

        close_btn = QPushButton("Exit")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

        self.setLayout(layout)

    def show_log_viewer(self):
        self.viewer = LogViewer(LOG_DIR)
        self.viewer.show()

    def launch_powerpoint(self):
        possible_paths = [
            os.path.join(os.environ.get("PROGRAMFILES", r"C:\\Program Files"), "Microsoft Office", "root", "Office16", "POWERPNT.EXE"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\\Program Files (x86)"), "Microsoft Office", "root", "Office16", "POWERPNT.EXE")
        ]

        for path in possible_paths:
            if os.path.exists(path):
                try:
                    subprocess.Popen([path])
                    return
                except Exception as e:
                    QMessageBox.critical(self, "Launch Error", f"Failed to launch PowerPoint:\n{str(e)}")
                    return

        QMessageBox.warning(self, "PowerPoint Not Found", "PowerPoint could not be launched automatically. Please open it manually.")

    def start_timer(self):
        duration = self.duration_input.text()
        if not duration.isdigit():
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number of minutes.")
            return

        self.presentation_minutes = int(duration)

        try:
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
            for _ in range(10):
                if self.ppt.Presentations.Count > 0:
                    break
                time.sleep(0.5)
            if self.ppt.Presentations.Count == 0:
                QMessageBox.warning(self, "PowerPoint", "Please open a presentation before starting.")
                return
            self.presentation = self.ppt.Presentations(1)
        except Exception as e:
            QMessageBox.critical(self, "PowerPoint Error", str(e))
            return

        self.start_time = time.time()
        self.slide_log = []
        self.overlay.show()
        self.overlay.move(200, 200)
        self.timer.start()

    def update_overlay(self):
        if not self.ppt or self.ppt.SlideShowWindows.Count == 0:
            self.end_session()
            return

        elapsed = time.time() - self.start_time
        total_time = self.presentation_minutes * 60
        remaining = max(0, total_time - elapsed)
        minutes = int(remaining // 60)
        seconds = int(remaining % 60)

        total_slides = self.presentation.Slides.Count
        current_slide = self.ppt.SlideShowWindows(1).View.CurrentShowPosition

        self.overlay.timer_label.setText(f"⏱️ Time left: {minutes}:{seconds:02}")
        self.overlay.slide_label.setText(f"Slide {current_slide} / {total_slides}")
        self.overlay.progress_bar.setValue(int((current_slide / total_slides) * 100))

        expected_slide = round((elapsed / total_time) * total_slides)
        expected_slide = max(1, min(expected_slide, total_slides))
        diff = current_slide - expected_slide

        if abs(diff) <= 3:
            pacing = "on track"
            self.overlay.message_label.setText("✅ You're on track!")
        elif diff > 3 and diff <= 6:
            pacing = "ahead"
            self.overlay.message_label.setText("🔵 You're ahead — consider slowing down.")
        elif diff > 6:
            pacing = "way ahead"
            self.overlay.message_label.setText("🔵 You're well ahead — consider pacing yourself.")
        elif diff < -3 and diff >= -6:
            pacing = "behind"
            self.overlay.message_label.setText("🟡 You're falling behind — pick up the pace.")
        else:
            pacing = "way behind"
            self.overlay.message_label.setText("🔴 You're well behind — consider skipping less critical slides.")

        self.slide_log.append({
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "elapsed_seconds": int(elapsed),
            "slide": current_slide,
            "pacing": pacing
        })

    def end_session(self):
        self.timer.stop()
        self.overlay.hide()

        if not self.slide_log:
            return

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filepath = os.path.join(LOG_DIR, f"log_{timestamp}.json")
        data = {
            "start_time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(self.start_time)),
            "duration_minutes": self.presentation_minutes,
            "slide_count": self.presentation.Slides.Count if self.presentation else 0,
            "entries": self.slide_log
        }
        with open(filepath, "w") as f:
            json.dump(data, f, indent=2)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())