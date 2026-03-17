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
    QDoubleSpinBox,
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
    QScrollArea,
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
    STACK_SIZE = 7
    MAX_POSITIONS = 25

    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("NERG Slide Capture")
        self.resize(1280, 820)

        default_drive_root = Path("/media/nerg/snailfish/gametophyte_imagery")
        local_fallback_root = Path.home() / "SlideCaptures"

        if default_drive_root.exists():
            self.output_root = default_drive_root
        else:
            self.output_root = local_fallback_root
            self.output_root.mkdir(parents=True, exist_ok=True)

        self.picam2 = None
        self.preview_widget = None
        self.capture_in_progress = False
        self.pending_file_path: Path | None = None
        self.pending_state: CaptureState | None = None
        self.pending_capture_kind = "stack"  # calibration | stack

        self.session_capture_count = 0
        self.calibration_captured = False
        self.last_stack_completed_name = ""

        self.stack_active = False
        self.stack_state: CaptureState | None = None
        self.stack_folder: Path | None = None
        self.stack_index = 0
        self.stack_started_at = ""

        if PI_CAMERA_AVAILABLE:
            try:
                self.picam2 = Picamera2()
                self.preview_config = self.picam2.create_preview_configuration(
                    main={"size": (1280, 720)}
                )
                self.still_config = self.picam2.create_still_configuration(
                    main={"size": (4056, 3040), "format": "RGB888"}
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
            self.reset_camera_defaults(show_status=False)

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_system_time_label)
        self.clock_timer.start(1000)
        self.update_system_time_label()

        QApplication.instance().installEventFilter(self)

        self.refresh_summary()
        self.update_stack_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        central.setLayout(main_layout)

        left_panel = QVBoxLayout()
        left_panel.setSpacing(6)
        main_layout.addLayout(left_panel, stretch=3)

        self.title_label = QLabel("NERG Slide Capture")
        self.title_label.setObjectName("TitleLabel")
        left_panel.addWidget(self.title_label)

        if self.picam2 is not None:
            self.preview_widget = QGlPicamera2(self.picam2, width=800, height=600)
            self.preview_widget.setMinimumSize(700, 520)
        else:
            self.preview_widget = QLabel("Camera preview\n(Test mode on this computer)")
            self.preview_widget.setAlignment(Qt.AlignCenter)
            self.preview_widget.setMinimumSize(700, 520)
            self.preview_widget.setObjectName("PreviewPlaceholder")

        left_panel.addWidget(self.preview_widget, stretch=1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_layout.addWidget(scroll, stretch=2)

        right_container = QWidget()
        scroll.setWidget(right_container)

        right_panel = QVBoxLayout()
        right_panel.setSpacing(10)
        right_container.setLayout(right_panel)

        session_group = QGroupBox("Session setup")
        session_layout = QFormLayout()
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

        self.camera_status_label = QLabel("Camera connected" if self.picam2 else "Test mode")
        self.camera_status_label.setObjectName("StatusChip")

        self.choose_output_root_btn = QPushButton("Change…")

        output_row = QHBoxLayout()
        output_row.addWidget(self.output_root_label, stretch=1)
        output_row.addWidget(self.choose_output_root_btn)

        session_layout.addRow("Date collected", self.collected_date_input)
        session_layout.addRow("Time collected", self.collected_time_input)
        session_layout.addRow("Site", self.site_code_input)
        session_layout.addRow("Output root", output_row)
        session_layout.addRow("System time", self.system_time_label)
        session_layout.addRow("Camera", self.camera_status_label)

        right_panel.addWidget(session_group)

        slide_group = QGroupBox("Slide workflow")
        slide_layout = QGridLayout()
        slide_group.setLayout(slide_layout)

        self.slide_number_spin = QSpinBox()
        self.slide_number_spin.setRange(0, 999)

        self.slide_position_spin = QSpinBox()
        self.slide_position_spin.setRange(1, self.MAX_POSITIONS)

        self.auto_increment_checkbox = QCheckBox("Auto advance to next replicate after completed stack")
        self.auto_increment_checkbox.setChecked(True)

        self.prev_position_btn = QPushButton("Previous replicate")
        self.next_position_btn = QPushButton("Next replicate")
        self.next_slide_btn = QPushButton("Next slide")

        slide_layout.addWidget(QLabel("Slide number"), 0, 0)
        slide_layout.addWidget(self.slide_number_spin, 0, 1)
        slide_layout.addWidget(QLabel("Replicate stack"), 1, 0)
        slide_layout.addWidget(self.slide_position_spin, 1, 1)
        slide_layout.addWidget(self.auto_increment_checkbox, 2, 0, 1, 2)
        slide_layout.addWidget(self.prev_position_btn, 3, 0)
        slide_layout.addWidget(self.next_position_btn, 3, 1)
        slide_layout.addWidget(self.next_slide_btn, 4, 0, 1, 2)

        right_panel.addWidget(slide_group)

        camera_group = QGroupBox("Camera settings")
        camera_layout = QFormLayout()
        camera_group.setLayout(camera_layout)

        self.manual_exposure_checkbox = QCheckBox("Use manual exposure + gain")
        self.shutter_spin = QSpinBox()
        self.shutter_spin.setRange(100, 500000)
        self.shutter_spin.setSingleStep(500)
        self.shutter_spin.setSuffix(" µs")
        self.shutter_spin.setValue(10000)

        self.gain_spin = QDoubleSpinBox()
        self.gain_spin.setRange(1.0, 16.0)
        self.gain_spin.setSingleStep(0.1)
        self.gain_spin.setDecimals(1)
        self.gain_spin.setValue(1.0)

        self.reset_camera_defaults_btn = QPushButton("Reset camera defaults")

        camera_layout.addRow(self.manual_exposure_checkbox)
        camera_layout.addRow("Shutter", self.shutter_spin)
        camera_layout.addRow("Gain", self.gain_spin)
        camera_layout.addRow(self.reset_camera_defaults_btn)

        right_panel.addWidget(camera_group)

        capture_group = QGroupBox("Capture workflow")
        capture_layout = QVBoxLayout()
        capture_group.setLayout(capture_layout)

        self.capture_calibration_btn = QPushButton("CAPTURE CALIBRATION IMAGE")
        self.capture_calibration_btn.setMinimumHeight(48)

        self.stack_btn = QPushButton("START STACK")
        self.stack_btn.setMinimumHeight(56)

        self.cancel_stack_btn = QPushButton("Cancel active stack")
        self.cancel_stack_btn.setMinimumHeight(42)

        self.calibration_status_label = QLabel("Calibration image: not captured")
        self.calibration_status_label.setObjectName("InfoLine")

        self.stack_status_label = QLabel("No active stack")
        self.stack_status_label.setObjectName("InfoLine")

        self.last_saved_label = QLabel("Last saved: none")
        self.last_saved_label.setObjectName("InfoLine")

        self.session_count_label = QLabel("Session images saved: 0")
        self.session_count_label.setObjectName("InfoLine")

        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setMinimumHeight(220)
        self.summary_box.setObjectName("SummaryBox")

        capture_layout.addWidget(self.capture_calibration_btn)
        capture_layout.addWidget(self.stack_btn)
        capture_layout.addWidget(self.cancel_stack_btn)
        capture_layout.addWidget(self.calibration_status_label)
        capture_layout.addWidget(self.stack_status_label)
        capture_layout.addWidget(self.last_saved_label)
        capture_layout.addWidget(self.session_count_label)
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
                padding: 0 0 2px 2px;
            }
            QGroupBox {
                background-color: white; border: 1px solid #dbe3ec; border-radius: 12px;
                margin-top: 10px; padding: 12px; font-weight: 700; color: #334155;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 10px; padding: 0 5px;
            }
            QLabel#MutedValueLabel { color: #475569; }
            QLabel#SystemTimeChip {
                background-color: #eef4ff; color: #1d4ed8; border: 1px solid #c7dbff;
                border-radius: 9px; padding: 4px 8px; font-weight: 700;
            }
            QLabel#StatusChip {
                border-radius: 9px; padding: 4px 8px; font-weight: 700;
                border: 1px solid #dbe3ec; background-color: #f8fafc; color: #334155;
            }
            QLabel#InfoLine { color: #475569; font-size: 12px; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QTextEdit {
                background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 9px; padding: 6px;
            }
            QPushButton {
                background-color: #e8eef7; border: 1px solid #c9d6e6; border-radius: 9px;
                padding: 8px 10px; font-weight: 600;
            }
            QPushButton:hover { background-color: #dce7f5; }
            QTextEdit#SummaryBox {
                background-color: #f8fafc; border: 1px solid #dbe3ec; border-radius: 10px;
            }
            QLabel#PreviewPlaceholder {
                background-color: #111827; color: #e5e7eb; font-size: 22px; font-weight: 600;
                border-radius: 16px; border: 1px solid #374151;
            }
            """
        )

        if self.picam2 is not None:
            self.camera_status_label.setStyleSheet(
                "QLabel { background-color: #ecfdf3; color: #15803d; border: 1px solid #bbf7d0; border-radius: 9px; padding: 4px 8px; font-weight: 700; }"
            )
        else:
            self.camera_status_label.setStyleSheet(
                "QLabel { background-color: #fff7ed; color: #c2410c; border: 1px solid #fed7aa; border-radius: 9px; padding: 4px 8px; font-weight: 700; }"
            )

    def _connect_signals(self) -> None:
        self.choose_output_root_btn.clicked.connect(self.choose_output_root)
        self.capture_calibration_btn.clicked.connect(self.capture_calibration_image)
        self.stack_btn.clicked.connect(self.handle_stack_button)
        self.cancel_stack_btn.clicked.connect(self.cancel_stack)

        self.prev_position_btn.clicked.connect(self.decrease_position)
        self.next_position_btn.clicked.connect(self.advance_position)
        self.next_slide_btn.clicked.connect(self.next_slide)

        self.collected_date_input.textChanged.connect(self.refresh_summary)
        self.collected_time_input.textChanged.connect(self.refresh_summary)
        self.site_code_input.textChanged.connect(self.refresh_summary)
        self.slide_number_spin.valueChanged.connect(self.refresh_summary)
        self.slide_position_spin.valueChanged.connect(self.refresh_summary)

        self.manual_exposure_checkbox.toggled.connect(self.apply_camera_controls)
        self.shutter_spin.valueChanged.connect(self.apply_camera_controls)
        self.gain_spin.valueChanged.connect(self.apply_camera_controls)
        self.reset_camera_defaults_btn.clicked.connect(self.reset_camera_defaults)

        if self.picam2 is not None and hasattr(self.preview_widget, "done_signal"):
            self.preview_widget.done_signal.connect(self.capture_done)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.modifiers() == Qt.NoModifier:
            if event.key() == Qt.Key_C:
                if self.stack_active:
                    self.handle_stack_button()
                return True
        return super().eventFilter(obj, event)

    def _load_defaults(self) -> None:
        now = datetime.now()
        self.collected_date_input.setText(now.strftime("%Y%m%d"))
        self.collected_time_input.setText(now.strftime("%H%M"))

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

    def current_site_folder(self, state: CaptureState | None = None) -> Path:
        state = state or self.get_state()
        return self.output_root / state.folder_name

    def build_calibration_filename(self, state: CaptureState) -> str:
        return f"{state.collected_date}_{state.collected_time}_{state.site_code}_CALIBRATION.png"

    def build_stack_folder_name(self, state: CaptureState) -> str:
        return f"{state.collected_date}_{state.collected_time}_{state.site_code}_S{state.slide_number:02d}_P{state.slide_position:03d}_STACK"

    def build_stack_frame_filename(self, state: CaptureState, z_index: int) -> str:
        return f"{state.collected_date}_{state.collected_time}_{state.site_code}_S{state.slide_number:02d}_P{state.slide_position:03d}_Z{z_index:02d}.png"

    def choose_output_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder", str(self.output_root))
        if folder:
            self.output_root = Path(folder)
            self.output_root_label.setText(folder)
            self.refresh_summary()

    def advance_position(self) -> None:
        self.slide_position_spin.setValue(min(self.MAX_POSITIONS, self.slide_position_spin.value() + 1))

    def decrease_position(self) -> None:
        self.slide_position_spin.setValue(max(1, self.slide_position_spin.value() - 1))

    def next_slide(self) -> None:
        if self.stack_active:
            QMessageBox.warning(self, "Active stack", "Finish or cancel the active stack before moving to the next slide.")
            return
        self.slide_number_spin.setValue(self.slide_number_spin.value() + 1)
        self.slide_position_spin.setValue(1)
        self.statusBar().showMessage(f"Moved to slide {self.slide_number_spin.value()}, replicate 1")

    def apply_camera_controls(self) -> None:
        self.refresh_summary()
        if self.picam2 is None:
            return
        try:
            if self.manual_exposure_checkbox.isChecked():
                self.picam2.set_controls({
                    "AeEnable": False,
                    "AwbEnable": False,
                    "ExposureTime": int(self.shutter_spin.value()),
                    "AnalogueGain": float(self.gain_spin.value()),
                })
                self.statusBar().showMessage(
                    f"Manual camera settings applied: shutter={self.shutter_spin.value()} µs, gain={self.gain_spin.value():.1f}"
                )
            else:
                self.picam2.set_controls({"AeEnable": True, "AwbEnable": True})
                self.statusBar().showMessage("Camera set to automatic exposure and white balance")
        except Exception as exc:
            self.statusBar().showMessage(f"Could not apply camera controls: {exc}")

    def reset_camera_defaults(self, show_status: bool = True) -> None:
        self.manual_exposure_checkbox.blockSignals(True)
        self.shutter_spin.blockSignals(True)
        self.gain_spin.blockSignals(True)

        self.manual_exposure_checkbox.setChecked(False)
        self.shutter_spin.setValue(10000)
        self.gain_spin.setValue(1.0)

        self.manual_exposure_checkbox.blockSignals(False)
        self.shutter_spin.blockSignals(False)
        self.gain_spin.blockSignals(False)

        self.refresh_summary()

        if self.picam2 is None:
            if show_status:
                self.statusBar().showMessage("Test mode: defaults reset in UI only")
            return

        try:
            self.picam2.set_controls({"AeEnable": True, "AwbEnable": True})
            if show_status:
                self.statusBar().showMessage("Camera defaults restored")
        except Exception as exc:
            self.statusBar().showMessage(f"Could not reset camera defaults: {exc}")

    def set_busy_ui(self, busy: bool) -> None:
        metadata_enabled = (not busy) and (not self.stack_active)
        self.capture_calibration_btn.setEnabled((not busy) and (not self.stack_active))
        self.stack_btn.setEnabled(not busy)
        self.cancel_stack_btn.setEnabled((not busy) and self.stack_active)
        self.prev_position_btn.setEnabled(metadata_enabled)
        self.next_position_btn.setEnabled(metadata_enabled)
        self.next_slide_btn.setEnabled(metadata_enabled)
        self.slide_number_spin.setEnabled(metadata_enabled)
        self.slide_position_spin.setEnabled(metadata_enabled)

    def update_stack_ui(self) -> None:
        if self.stack_active:
            self.stack_btn.setText(f"CAPTURE Z{self.stack_index + 1:02d} OF {self.STACK_SIZE}")
            self.stack_status_label.setText(
                f"Active stack: replicate {self.stack_state.slide_position if self.stack_state else '?'} | next plane {self.stack_index + 1} of {self.STACK_SIZE}"
            )
        else:
            self.stack_btn.setText("START STACK")
            self.stack_status_label.setText("No active stack")

        self.calibration_status_label.setText(
            "Calibration image: captured" if self.calibration_captured else "Calibration image: not captured"
        )
        self.session_count_label.setText(f"Session images saved: {self.session_capture_count}")

    def refresh_summary(self) -> None:
        state = self.get_state()
        error = self.validate_state(state)
        folder = self.current_site_folder(state)
        calibration_file = folder / self.build_calibration_filename(state) if error is None else "Invalid metadata"
        next_stack = folder / self.build_stack_folder_name(state) if error is None else "Invalid metadata"

        text = (
            f"Session folder:\n{folder}\n\n"
            f"Calibration image:\n{calibration_file}\n\n"
            f"Next stack folder:\n{next_stack}\n\n"
            f"Slide: {state.slide_number}\n"
            f"Replicate stack: {state.slide_position} of {self.MAX_POSITIONS}\n\n"
            f"Camera mode: {'Manual' if self.manual_exposure_checkbox.isChecked() else 'Auto'}\n"
            f"Shutter: {self.shutter_spin.value()} µs\n"
            f"Gain: {self.gain_spin.value():.1f}\n\n"
            "Workflow:\n"
            "1. Capture one calibration image.\n"
            "2. Start a stack.\n"
            "3. Press c seven times to capture Z01–Z07.\n"
            "4. After Z07, the stack completes automatically.\n\n"
            "ImageJ/Fiji:\n"
            "File → Import → Image Sequence\n"
            "Select one *_STACK folder to load Z01–Z07."
        )
        if error is not None:
            text += f"\n\nWarning:\n{error}"

        self.summary_box.setPlainText(text)
        self.update_stack_ui()

    def begin_capture(self, full_path: Path, state: CaptureState, kind: str, status_text: str) -> None:
        self.pending_file_path = full_path
        self.pending_state = state
        self.pending_capture_kind = kind
        self.capture_in_progress = True
        self.set_busy_ui(True)
        self.statusBar().showMessage(status_text)

        try:
            if self.picam2 is None:
                full_path.parent.mkdir(parents=True, exist_ok=True)
                with open(full_path, "w", encoding="utf-8") as handle:
                    handle.write(f"dummy {kind} image")
                self.capture_done(None)
            else:
                full_path.parent.mkdir(parents=True, exist_ok=True)
                self.picam2.switch_mode_and_capture_file(
                    self.still_config,
                    str(full_path),
                    signal_function=self.preview_widget.signal_done,
                )
        except Exception as exc:
            self.capture_in_progress = False
            self.set_busy_ui(False)
            QMessageBox.critical(self, "Capture failed", f"Could not capture image:\n\n{exc}")

    def capture_calibration_image(self) -> None:
        if self.capture_in_progress or self.stack_active:
            return

        state = self.get_state()
        error = self.validate_state(state)
        if error is not None:
            QMessageBox.warning(self, "Check metadata", error)
            return

        folder = self.current_site_folder(state)
        full_path = folder / self.build_calibration_filename(state)
        self.begin_capture(full_path, state, "calibration", f"Capturing calibration image {full_path.name}...")

    def handle_stack_button(self) -> None:
        if self.capture_in_progress:
            return
        if not self.stack_active:
            self.start_stack()
        else:
            self.capture_stack_plane()

    def start_stack(self) -> None:
        state = self.get_state()
        error = self.validate_state(state)
        if error is not None:
            QMessageBox.warning(self, "Check metadata", error)
            return

        stack_folder = self.current_site_folder(state) / self.build_stack_folder_name(state)
        if stack_folder.exists():
            QMessageBox.warning(
                self,
                "Stack already exists",
                f"A stack folder already exists for this slide and replicate:\n\n{stack_folder.name}"
            )
            return

        stack_folder.mkdir(parents=True, exist_ok=False)

        self.stack_active = True
        self.stack_state = state
        self.stack_folder = stack_folder
        self.stack_index = 0
        self.stack_started_at = datetime.now().isoformat(timespec="seconds")

        self.refresh_summary()
        self.set_busy_ui(False)
        self.statusBar().showMessage(
            f"Stack started for slide {state.slide_number}, replicate {state.slide_position}. Press c seven times to capture Z01–Z07."
        )

    def cancel_stack(self) -> None:
        if not self.stack_active or self.capture_in_progress:
            return

        if self.stack_folder is not None and self.stack_folder.exists():
            for child in self.stack_folder.iterdir():
                child.unlink()
            self.stack_folder.rmdir()

        self.stack_active = False
        self.stack_state = None
        self.stack_folder = None
        self.stack_index = 0
        self.stack_started_at = ""

        self.refresh_summary()
        self.set_busy_ui(False)
        self.statusBar().showMessage("Stack canceled")

    def capture_stack_plane(self) -> None:
        if not self.stack_active or self.stack_state is None or self.stack_folder is None:
            return

        z_index = self.stack_index + 1
        full_path = self.stack_folder / self.build_stack_frame_filename(self.stack_state, z_index)
        self.begin_capture(full_path, self.stack_state, "stack", f"Capturing Z{z_index:02d} of {self.STACK_SIZE}...")

    def capture_done(self, job) -> None:
        try:
            if self.picam2 is not None and job is not None:
                self.picam2.wait(job)

            if self.pending_file_path is None or self.pending_state is None:
                return

            self.session_capture_count += 1
            self.last_saved_label.setText(f"Last saved: {self.pending_file_path.name}")

            if self.pending_capture_kind == "calibration":
                self.calibration_captured = True
                self.write_calibration_log(self.pending_file_path, self.pending_state)
                self.statusBar().showMessage(f"Saved calibration image {self.pending_file_path.name}")

            elif self.pending_capture_kind == "stack":
                z_index = self.stack_index + 1
                self.write_stack_frame_log(self.pending_file_path, self.pending_state, z_index)
                self.statusBar().showMessage(f"Saved Z{z_index:02d} of {self.STACK_SIZE}")
                self.stack_index += 1

                if self.stack_index >= self.STACK_SIZE:
                    self.finish_stack()

            self.capture_in_progress = False
            self.pending_file_path = None
            self.pending_state = None
            self.pending_capture_kind = "stack"
            self.set_busy_ui(False)
            self.refresh_summary()

        except Exception as exc:
            QMessageBox.critical(self, "Capture completion failed", f"{exc}")
            self.capture_in_progress = False
            self.pending_file_path = None
            self.pending_state = None
            self.pending_capture_kind = "stack"
            self.set_busy_ui(False)
            self.refresh_summary()

    def finish_stack(self) -> None:
        if self.stack_state is None or self.stack_folder is None:
            return

        self.write_stack_manifest(self.stack_folder, self.stack_state)
        self.write_site_stack_log(self.stack_folder, self.stack_state)

        self.statusBar().showMessage(
            f"Completed stack for slide {self.stack_state.slide_number}, replicate {self.stack_state.slide_position}"
        )
        self.last_saved_label.setText(f"Last saved: {self.stack_folder.name}")
        self.stack_active = False
        self.last_stack_completed_name = self.stack_folder.name

        if self.auto_increment_checkbox.isChecked():
            self.advance_position()

        self.stack_state = None
        self.stack_folder = None
        self.stack_index = 0
        self.stack_started_at = ""

    def write_calibration_log(self, path: Path, state: CaptureState) -> None:
        log_file = path.parent / "calibration_log.csv"
        exists = log_file.exists()
        with open(log_file, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at", "site_folder", "site_code", "slide_number",
                    "filename", "full_path"
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                state.folder_name, state.site_code, state.slide_number,
                path.name, str(path)
            ])

    def write_stack_frame_log(self, path: Path, state: CaptureState, z_index: int) -> None:
        if self.stack_folder is None:
            return
        log_file = self.stack_folder / "stack_frame_log.csv"
        exists = log_file.exists()
        with open(log_file, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at", "site_folder", "site_code", "slide_number",
                    "slide_position", "stack_folder", "z_index", "filename", "full_path"
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                state.folder_name, state.site_code, state.slide_number,
                state.slide_position, self.stack_folder.name, z_index,
                path.name, str(path)
            ])

    def write_stack_manifest(self, folder: Path, state: CaptureState) -> None:
        manifest_file = folder / "stack_manifest.csv"
        with open(manifest_file, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow([
                "saved_at", "started_at", "site_folder", "site_code",
                "slide_number", "slide_position", "stack_folder", "frame_count"
            ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                self.stack_started_at, state.folder_name, state.site_code,
                state.slide_number, state.slide_position, folder.name, self.STACK_SIZE
            ])

    def write_site_stack_log(self, folder: Path, state: CaptureState) -> None:
        site_log = folder.parent / "stack_log.csv"
        exists = site_log.exists()
        with open(site_log, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at", "started_at", "site_folder", "site_code",
                    "slide_number", "slide_position", "stack_folder",
                    "frame_count", "stack_path"
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                self.stack_started_at, state.folder_name, state.site_code,
                state.slide_number, state.slide_position, folder.name,
                self.STACK_SIZE, str(folder)
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