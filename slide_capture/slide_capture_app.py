#!/usr/bin/env python3
from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, QEvent
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

PI_CAMERA_AVAILABLE = True

try:
    from picamera2 import Picamera2
    from picamera2.previews.qt import QGlPicamera2
except Exception:
    PI_CAMERA_AVAILABLE = False


SITE_FOLDER_PATTERN = re.compile(r"^\d{8}_\d{4}_[A-Za-z0-9_-]+$")
SITE_CODE_PATTERN = re.compile(r"^[A-Za-z0-9_-]+$")


@dataclass
class CaptureState:
    collected_date: str = ""
    collected_time: str = ""
    site_code: str = ""
    slide_number: int = 0
    slide_position: int = 1

    @property
    def folder_name(self) -> str:
        return f"{self.collected_date}_{self.collected_time}_{self.site_code}"


class SlideCaptureApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("NERG Slide Capture")
        self.resize(1220, 760)

        #self.output_root = Path.home() / "SlideCaptures"
        #self.output_root.mkdir(parents=True, exist_ok=True)
        
        default_drive_root = Path("media/nerg/snailfish/gametophyte_imagery")
        local_fallback_root = Path.home() / "SlideCaptures"
        
        if default_drive_root.exists():
            self.output_root = default_drive_root
        else:
            self.output_root = local_fallback_root
            self.output_root.mkdir(parents=True, exist_ok=True)
            

        self.picam2 = None
        self.preview_widget = None
        self.capture_count = 0
        self.capture_in_progress = False
        self.pending_file_path: Path | None = None
        self.pending_state: CaptureState | None = None

        if PI_CAMERA_AVAILABLE:
            try:
                self.picam2 = Picamera2()
                self.preview_config = self.picam2.create_preview_configuration(
                    main={"size": (1280, 720)}
                )
                self.still_config = self.picam2.create_still_configuration(
                    main={"size": (4056, 3040)}
                )
                self.picam2.configure(self.preview_config)
            except Exception:
                self.picam2 = None

        self._build_ui()
        self.apply_styles()
        self._connect_signals()
        self._load_defaults()

        if self.picam2 is not None:
            self.picam2.start()

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_system_time_label)
        self.clock_timer.start(1000)
        self.update_system_time_label()

        self.status_flash_timer = QTimer(self)
        self.status_flash_timer.setSingleShot(True)
        self.status_flash_timer.timeout.connect(self.reset_capture_button_style)

        QApplication.instance().installEventFilter(self)

        self.refresh_summary()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)
        central.setLayout(main_layout)

        left_panel = QVBoxLayout()
        left_panel.setSpacing(4)
        main_layout.addLayout(left_panel, stretch=3)

        self.title_label = QLabel("NERG Slide Capture")
        self.title_label.setObjectName("TitleLabel")
        left_panel.addWidget(self.title_label)

        if self.picam2 is not None:
            self.preview_widget = QGlPicamera2(self.picam2, width=800, height=600)
            self.preview_widget.setMinimumSize(700, 500)
        else:
            self.preview_widget = QLabel("Camera preview\n(Test mode on this computer)")
            self.preview_widget.setAlignment(Qt.AlignCenter)
            self.preview_widget.setMinimumSize(700, 500)
            self.preview_widget.setObjectName("PreviewPlaceholder")

        left_panel.addWidget(self.preview_widget)

        right_panel = QVBoxLayout()
        right_panel.setSpacing(12)
        main_layout.addLayout(right_panel, stretch=2)

        session_group = QGroupBox("Session setup")
        session_layout = QFormLayout()
        session_layout.setSpacing(10)
        session_group.setLayout(session_layout)

        self.collected_date_input = QLineEdit()
        self.collected_date_input.setPlaceholderText("YYYYMMDD")

        self.collected_time_input = QLineEdit()
        self.collected_time_input.setPlaceholderText("HHMM")

        self.site_code_input = QLineEdit()
        self.site_code_input.setPlaceholderText("SITE")

        self.output_root_label = QLabel(str(self.output_root))
        self.output_root_label.setWordWrap(True)
        self.output_root_label.setObjectName("MutedValueLabel")

        self.system_time_label = QLabel("")
        self.system_time_label.setObjectName("SystemTimeChip")

        self.camera_status_label = QLabel("")
        self.camera_status_label.setObjectName("StatusChip")
        self.camera_status_label.setText("Camera connected" if self.picam2 else "Test mode")

        self.choose_output_root_btn = QPushButton("Change…")

        output_row = QHBoxLayout()
        output_row.setSpacing(8)
        output_row.addWidget(self.output_root_label, stretch=1)
        output_row.addWidget(self.choose_output_root_btn)

        session_layout.addRow("Date collected", self.collected_date_input)
        session_layout.addRow("Time collected", self.collected_time_input)
        session_layout.addRow("Site", self.site_code_input)
        session_layout.addRow("Output root", output_row)
        session_layout.addRow("System time", self.system_time_label)
        session_layout.addRow("Camera", self.camera_status_label)

        right_panel.addWidget(session_group)

        slide_group = QGroupBox("Slide controls")
        slide_layout = QGridLayout()
        slide_layout.setHorizontalSpacing(10)
        slide_layout.setVerticalSpacing(10)
        slide_group.setLayout(slide_layout)

        self.slide_number_spin = QSpinBox()
        self.slide_number_spin.setRange(0, 999)

        self.slide_position_spin = QSpinBox()
        self.slide_position_spin.setRange(1, 999)

        self.auto_increment_checkbox = QCheckBox("Auto advance position after capture")
        self.auto_increment_checkbox.setChecked(True)

        self.prev_position_btn = QPushButton("Previous position")
        self.next_position_btn = QPushButton("Next position")
        self.next_slide_btn = QPushButton("Next slide")

        slide_layout.addWidget(QLabel("Slide number"), 0, 0)
        slide_layout.addWidget(self.slide_number_spin, 0, 1)

        slide_layout.addWidget(QLabel("Slide position"), 1, 0)
        slide_layout.addWidget(self.slide_position_spin, 1, 1)

        slide_layout.addWidget(self.auto_increment_checkbox, 2, 0, 1, 2)
        slide_layout.addWidget(self.prev_position_btn, 3, 0)
        slide_layout.addWidget(self.next_position_btn, 3, 1)
        slide_layout.addWidget(self.next_slide_btn, 4, 0, 1, 2)

        right_panel.addWidget(slide_group)

        capture_group = QGroupBox("Capture")
        capture_layout = QVBoxLayout()
        capture_layout.setSpacing(10)
        capture_group.setLayout(capture_layout)

        self.capture_btn = QPushButton("CAPTURE IMAGE")
        self.capture_btn.setObjectName("CaptureButton")
        self.capture_btn.setMinimumHeight(70)

        self.last_file_label = QLabel("Last file: none")
        self.last_file_label.setObjectName("InfoLine")

        self.capture_count_label = QLabel("Session captures: 0")
        self.capture_count_label.setObjectName("InfoLine")

        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setObjectName("SummaryBox")

        capture_layout.addWidget(self.capture_btn)
        capture_layout.addWidget(self.last_file_label)
        capture_layout.addWidget(self.capture_count_label)
        capture_layout.addWidget(self.summary_box)

        right_panel.addWidget(capture_group)
        right_panel.addStretch(1)

        self.setStatusBar(QStatusBar())

    def apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background-color: #f4f7fb; }
            QWidget { font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #1f2937; }
            QLabel#TitleLabel {
                font-size: 18px; font-weight: 700; color: #0f172a;
                margin: 0px; padding: 0px 0px 2px 2px; max-height: 24px;
            }
            QGroupBox {
                background-color: white; border: 1px solid #dbe3ec; border-radius: 14px;
                margin-top: 10px; padding: 12px; font-weight: 700; font-size: 14px; color: #334155;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 12px; padding: 0 6px 0 6px; color: #334155;
            }
            QLabel { color: #334155; }
            QLabel#MutedValueLabel { color: #475569; padding: 2px 0px 2px 0px; }
            QLabel#SystemTimeChip {
                background-color: #eef4ff; color: #1d4ed8; border: 1px solid #c7dbff;
                border-radius: 10px; padding: 5px 9px; font-weight: 700;
            }
            QLabel#StatusChip {
                border-radius: 10px; padding: 5px 9px; font-weight: 700;
                border: 1px solid #dbe3ec; background-color: #f8fafc; color: #334155;
            }
            QLabel#InfoLine { color: #475569; font-size: 12px; }
            QLineEdit, QSpinBox, QTextEdit {
                background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 10px; padding: 7px;
            }
            QLineEdit:focus, QSpinBox:focus, QTextEdit:focus {
                border: 2px solid #4f86f7; background-color: white;
            }
            QPushButton {
                background-color: #e8eef7; border: 1px solid #c9d6e6; border-radius: 10px;
                padding: 9px 12px; font-weight: 600;
            }
            QPushButton:hover { background-color: #dce7f5; }
            QPushButton:pressed { background-color: #cfdcf0; }
            QPushButton#CaptureButton {
                background-color: #2563eb; color: white; border: none; border-radius: 14px;
                font-size: 18px; font-weight: 700; padding: 12px;
            }
            QPushButton#CaptureButton:hover { background-color: #1d4ed8; }
            QPushButton#CaptureButton:pressed { background-color: #1e40af; }
            QTextEdit#SummaryBox {
                background-color: #f8fafc; border: 1px solid #dbe3ec; border-radius: 12px;
                padding: 8px; color: #334155;
            }
            QLabel#PreviewPlaceholder {
                background-color: #111827; color: #e5e7eb; font-size: 22px; font-weight: 600;
                border-radius: 18px; border: 1px solid #374151;
            }
            QCheckBox { spacing: 8px; color: #334155; }
            QStatusBar { background: white; border-top: 1px solid #dbe3ec; color: #475569; }
            """
        )

        if self.picam2 is not None:
            self.camera_status_label.setStyleSheet(
                "QLabel { background-color: #ecfdf3; color: #15803d; border: 1px solid #bbf7d0; border-radius: 10px; padding: 5px 9px; font-weight: 700; }"
            )
        else:
            self.camera_status_label.setStyleSheet(
                "QLabel { background-color: #fff7ed; color: #c2410c; border: 1px solid #fed7aa; border-radius: 10px; padding: 5px 9px; font-weight: 700; }"
            )

    def flash_capture_success(self) -> None:
        self.capture_btn.setStyleSheet(
            "QPushButton { background-color: #16a34a; color: white; border: none; border-radius: 14px; font-size: 18px; font-weight: 700; padding: 12px; }"
        )
        self.status_flash_timer.start(220)

    def reset_capture_button_style(self) -> None:
        self.capture_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #2563eb; color: white; border: none; border-radius: 14px;
                font-size: 18px; font-weight: 700; padding: 12px;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
            """
        )

    def _connect_signals(self) -> None:
        self.choose_output_root_btn.clicked.connect(self.choose_output_root)
        self.capture_btn.clicked.connect(self.capture_image)

        self.prev_position_btn.clicked.connect(self.decrease_position)
        self.next_position_btn.clicked.connect(self.advance_position)
        self.next_slide_btn.clicked.connect(self.next_slide)

        self.collected_date_input.textChanged.connect(self.refresh_summary)
        self.collected_time_input.textChanged.connect(self.refresh_summary)
        self.site_code_input.textChanged.connect(self.refresh_summary)
        self.slide_number_spin.valueChanged.connect(self.refresh_summary)
        self.slide_position_spin.valueChanged.connect(self.refresh_summary)

        self.site_code_input.returnPressed.connect(self.capture_btn.setFocus)

        if self.picam2 is not None and hasattr(self.preview_widget, "done_signal"):
            self.preview_widget.done_signal.connect(self.capture_done)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.modifiers() == Qt.NoModifier:
            if event.key() == Qt.Key_C:
                self.capture_image()
                return True
            if event.key() == Qt.Key_N:
                self.next_slide()
                return True
            if event.key() == Qt.Key_Right:
                self.advance_position()
                return True
            if event.key() == Qt.Key_Left:
                self.decrease_position()
                return True

        return super().eventFilter(obj, event)

    def _load_defaults(self) -> None:
        now = datetime.now()
        self.collected_date_input.setText(now.strftime("%Y%m%d"))
        self.collected_time_input.setText(now.strftime("%H%M"))
        self.capture_btn.setFocus()

    def update_system_time_label(self) -> None:
        self.system_time_label.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def get_state(self) -> CaptureState:
        return CaptureState(
            collected_date=self.collected_date_input.text().strip(),
            collected_time=self.collected_time_input.text().strip(),
            site_code=self.site_code_input.text().strip(),
            slide_number=self.slide_number_spin.value(),
            slide_position=self.slide_position_spin.value(),
        )

    def validate_state(self, state: CaptureState) -> str | None:
        if not re.fullmatch(r"\d{8}", state.collected_date):
            return "Date collected must be 8 digits in YYYYMMDD format."
        if not re.fullmatch(r"\d{4}", state.collected_time):
            return "Time collected must be 4 digits in HHMM format."
        if not SITE_CODE_PATTERN.fullmatch(state.site_code):
            return "Site must use only letters, numbers, underscores, or hyphens."
        if not SITE_FOLDER_PATTERN.fullmatch(state.folder_name):
            return "Folder name must match DATECOLLECTED_TIMECOLLECTED_SITE."
        return None

    def current_site_folder(self) -> Path:
        return self.output_root / self.get_state().folder_name

    def build_filename(self, state: CaptureState) -> str:
        now = datetime.now()
        return f"{now.strftime('%Y%m%d')}_{now.strftime('%H%M')}_{state.site_code}_S{state.slide_number:02d}_P{state.slide_position:03d}.jpg"

    def refresh_summary(self) -> None:
        state = self.get_state()
        error = self.validate_state(state)
        folder = self.output_root / state.folder_name
        next_file = self.build_filename(state) if error is None else "Invalid metadata"

        text = (
            f"Site folder:\n{folder}\n\n"
            f"Next file:\n{next_file}\n\n"
            f"Slide: {state.slide_number}\n"
            f"Position: {state.slide_position}\n\n"
            "Shortcuts:\n"
            "c = capture\n"
            "n = next slide\n"
            "← / → = change position\n\n"
            "Processed timestamps use system time."
        )
        if error is not None:
            text += f"\n\nWarning:\n{error}"
        self.summary_box.setPlainText(text)

    def choose_output_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder", str(self.output_root))
        if folder:
            self.output_root = Path(folder)
            self.output_root_label.setText(folder)
            self.refresh_summary()

    def advance_position(self) -> None:
        self.slide_position_spin.setValue(self.slide_position_spin.value() + 1)

    def decrease_position(self) -> None:
        self.slide_position_spin.setValue(max(1, self.slide_position_spin.value() - 1))

    def next_slide(self) -> None:
        self.slide_number_spin.setValue(self.slide_number_spin.value() + 1)
        self.slide_position_spin.setValue(1)
        self.statusBar().showMessage(f"Moved to slide {self.slide_number_spin.value()}, position 1")

    def capture_image(self) -> None:
        if self.capture_in_progress:
            return

        state = self.get_state()
        error = self.validate_state(state)
        if error is not None:
            QMessageBox.warning(self, "Check metadata", error)
            return

        folder = self.output_root / state.folder_name
        folder.mkdir(parents=True, exist_ok=True)

        filename = self.build_filename(state)
        full_path = folder / filename

        self.pending_file_path = full_path
        self.pending_state = state
        self.capture_in_progress = True
        self.capture_btn.setEnabled(False)
        self.statusBar().showMessage(f"Capturing {filename}...")

        try:
            if self.picam2 is None:
                with open(full_path, "w", encoding="utf-8") as handle:
                    handle.write("dummy image for testing GUI")
                self.capture_done(None)
            else:
                self.picam2.switch_mode_and_capture_file(
                    self.still_config,
                    str(full_path),
                    signal_function=self.preview_widget.signal_done,
                )
        except Exception as exc:
            self.capture_in_progress = False
            self.capture_btn.setEnabled(True)
            QMessageBox.critical(self, "Capture failed", f"Could not capture image:\n\n{exc}")

    def capture_done(self, job) -> None:
        try:
            if self.picam2 is not None and job is not None:
                self.picam2.wait(job)

            if self.pending_file_path is None or self.pending_state is None:
                return

            self.capture_count += 1
            self.last_file_label.setText(f"Last file: {self.pending_file_path.name}")
            self.capture_count_label.setText(f"Session captures: {self.capture_count}")
            self.write_log(self.pending_file_path, self.pending_state)
            self.flash_capture_success()
            self.statusBar().showMessage(f"Saved {self.pending_file_path.name}")

            if self.auto_increment_checkbox.isChecked():
                self.advance_position()

        except Exception as exc:
            QMessageBox.critical(self, "Capture completion failed", f"{exc}")
        finally:
            self.capture_in_progress = False
            self.capture_btn.setEnabled(True)
            self.pending_file_path = None
            self.pending_state = None
            self.refresh_summary()

    def write_log(self, path: Path, state: CaptureState) -> None:
        log_file = path.parent / "capture_log.csv"
        exists = log_file.exists()

        with open(log_file, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at",
                    "site_folder",
                    "site_code",
                    "slide_number",
                    "slide_position",
                    "filename",
                    "full_path",
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                state.folder_name,
                state.site_code,
                state.slide_number,
                state.slide_position,
                path.name,
                str(path),
            ])

    def closeEvent(self, event) -> None:
        try:
            if self.picam2 is not None:
                self.picam2.stop()
        except Exception:
            pass
        event.accept()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("NERG Slide Capture")
    window = SlideCaptureApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
