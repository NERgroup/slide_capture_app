#!/usr/bin/env python3
from __future__ import annotations

import csv
import math
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
    slide_number: int = 1
    replicate_number: int = 1

    @property
    def folder_name(self) -> str:
        return f"{self.collected_date}_{self.collected_time}_{self.site_code}"


class SlideCaptureApp(QMainWindow):
    X_START = 143
    X_END = 130
    Y_START = 10
    Y_END = 20

    DEFAULT_PHOTOS_PER_SLIDE = 50
    MIN_PHOTOS_PER_SLIDE = 1
    MAX_PHOTOS_PER_SLIDE = 400

    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("NERG Slide Capture")
        self.resize(1340, 860)

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
        self.pending_capture_kind = "replicate"  # calibration | replicate
        self.pending_capture_coord: tuple[int, int] | None = None

        self.session_capture_count = 0

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

        self.grid_coords: list[tuple[int, int]] = []
        self.rebuild_grid()

        if self.picam2 is not None:
            self.picam2.start()
            self.reset_camera_defaults(show_status=False)

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_system_time_label)
        self.clock_timer.start(1000)
        self.update_system_time_label()

        QApplication.instance().installEventFilter(self)

        self.refresh_summary()
        self.update_capture_ui()

    # ----------------------------
    # Dynamic grid helpers
    # ----------------------------

    def photos_per_slide(self) -> int:
        return self.photos_per_slide_spin.value()

    def max_replicates(self) -> int:
        return self.photos_per_slide()

    def grid_shape_for_count(self, n: int) -> tuple[int, int]:
        """
        Choose rows/cols that:
        - cover at least n points
        - roughly match the physical aspect ratio of the capture area
        - keep the grid as compact as possible
        """
        if n <= 1:
            return 1, 1

        x_span = abs(self.X_START - self.X_END)
        y_span = abs(self.Y_END - self.Y_START)
        aspect = x_span / y_span if y_span != 0 else 1.0

        best_rows = 1
        best_cols = n
        best_score = None

        for rows in range(1, n + 1):
            cols = math.ceil(n / rows)

            occupancy_penalty = (rows * cols) - n
            shape_penalty = abs((cols / rows) - aspect)

            score = (occupancy_penalty * 10.0) + shape_penalty

            if best_score is None or score < best_score:
                best_score = score
                best_rows = rows
                best_cols = cols

            if rows > math.sqrt(n) * 3:
                # no need to search absurdly tall grids
                break

        return best_rows, best_cols

    def build_axis(self, start: int, end: int, n: int) -> list[int]:
        if n <= 1:
            return [int(round(start))]
        step = (end - start) / (n - 1)
        return [int(round(start + i * step)) for i in range(n)]

    def build_upc_grid(self, n_points: int) -> list[tuple[int, int]]:
        rows, cols = self.grid_shape_for_count(n_points)
        xs = self.build_axis(self.X_START, self.X_END, cols)
        ys = self.build_axis(self.Y_START, self.Y_END, rows)

        coords: list[tuple[int, int]] = []
        for row_idx, y in enumerate(ys):
            x_row = xs if row_idx % 2 == 0 else list(reversed(xs))
            for x in x_row:
                coords.append((x, y))

        return coords[:n_points]

    def rebuild_grid(self) -> None:
        self.grid_coords = self.build_upc_grid(self.photos_per_slide())
        self.replicate_spin.setMaximum(self.max_replicates())

        if self.replicate_spin.value() > self.max_replicates():
            self.replicate_spin.setValue(self.max_replicates())

    def coord_for_replicate(self, replicate_number: int) -> tuple[int, int]:
        if not self.grid_coords:
            self.rebuild_grid()

        idx = max(1, min(replicate_number, self.max_replicates())) - 1
        return self.grid_coords[idx]

    def coord_text(self, replicate_number: int) -> str:
        x, y = self.coord_for_replicate(replicate_number)
        return f"{x}x{y}"

    # ----------------------------
    # UI
    # ----------------------------

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)
        central.setLayout(main_layout)

        left_panel = QVBoxLayout()
        left_panel.setSpacing(8)
        main_layout.addLayout(left_panel, stretch=3)

        self.title_label = QLabel("NERG Slide Capture")
        self.title_label.setObjectName("TitleLabel")
        left_panel.addWidget(self.title_label)

        self.subtitle_label = QLabel(
            "Single calibration image per site • configurable replicate photos per slide • scaled snake grid"
        )
        self.subtitle_label.setObjectName("SubtitleLabel")
        left_panel.addWidget(self.subtitle_label)

        if self.picam2 is not None:
            self.preview_widget = QGlPicamera2(self.picam2, width=900, height=620)
            self.preview_widget.setMinimumSize(760, 560)
        else:
            self.preview_widget = QLabel("Camera preview\n(Test mode on this computer)")
            self.preview_widget.setAlignment(Qt.AlignCenter)
            self.preview_widget.setMinimumSize(760, 560)
            self.preview_widget.setObjectName("PreviewPlaceholder")

        left_panel.addWidget(self.preview_widget, stretch=1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFrameShape(QScrollArea.NoFrame)
        main_layout.addWidget(scroll, stretch=2)

        right_container = QWidget()
        scroll.setWidget(right_container)

        right_panel = QVBoxLayout()
        right_panel.setSpacing(12)
        right_container.setLayout(right_panel)

        # ---------------- Session setup ----------------
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

        # ---------------- Slide workflow ----------------
        slide_group = QGroupBox("Slide workflow")
        slide_layout = QGridLayout()
        slide_layout.setHorizontalSpacing(10)
        slide_layout.setVerticalSpacing(10)
        slide_group.setLayout(slide_layout)

        self.total_slides_spin = QSpinBox()
        self.total_slides_spin.setRange(1, 999)
        self.total_slides_spin.setValue(1)

        self.slide_number_spin = QSpinBox()
        self.slide_number_spin.setRange(1, 999)
        self.slide_number_spin.setValue(1)

        self.photos_per_slide_spin = QSpinBox()
        self.photos_per_slide_spin.setRange(self.MIN_PHOTOS_PER_SLIDE, self.MAX_PHOTOS_PER_SLIDE)
        self.photos_per_slide_spin.setValue(self.DEFAULT_PHOTOS_PER_SLIDE)

        self.replicate_spin = QSpinBox()
        self.replicate_spin.setRange(1, self.DEFAULT_PHOTOS_PER_SLIDE)
        self.replicate_spin.setValue(1)

        self.auto_increment_checkbox = QCheckBox("Auto advance after each captured photo")
        self.auto_increment_checkbox.setChecked(True)

        self.prev_replicate_btn = QPushButton("Previous replicate")
        self.next_replicate_btn = QPushButton("Next replicate")
        self.next_slide_btn = QPushButton("Next slide")

        slide_layout.addWidget(QLabel("Slides in site"), 0, 0)
        slide_layout.addWidget(self.total_slides_spin, 0, 1)
        slide_layout.addWidget(QLabel("Current slide"), 1, 0)
        slide_layout.addWidget(self.slide_number_spin, 1, 1)
        slide_layout.addWidget(QLabel("Photos per slide"), 2, 0)
        slide_layout.addWidget(self.photos_per_slide_spin, 2, 1)
        slide_layout.addWidget(QLabel("Replicate photo"), 3, 0)
        slide_layout.addWidget(self.replicate_spin, 3, 1)
        slide_layout.addWidget(self.auto_increment_checkbox, 4, 0, 1, 2)
        slide_layout.addWidget(self.prev_replicate_btn, 5, 0)
        slide_layout.addWidget(self.next_replicate_btn, 5, 1)
        slide_layout.addWidget(self.next_slide_btn, 6, 0, 1, 2)

        right_panel.addWidget(slide_group)

        # ---------------- Camera settings ----------------
        camera_group = QGroupBox("Camera settings")
        camera_layout = QFormLayout()
        camera_layout.setSpacing(10)
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

        # ---------------- Capture workflow ----------------
        capture_group = QGroupBox("Capture workflow")
        capture_layout = QVBoxLayout()
        capture_layout.setSpacing(10)
        capture_group.setLayout(capture_layout)

        self.capture_calibration_btn = QPushButton("CAPTURE SITE CALIBRATION")
        self.capture_calibration_btn.setMinimumHeight(50)
        self.capture_calibration_btn.setObjectName("PrimaryButton")

        self.capture_btn = QPushButton("CAPTURE 143x10")
        self.capture_btn.setMinimumHeight(58)
        self.capture_btn.setObjectName("AccentButton")

        indicator_grid = QGridLayout()
        indicator_grid.setHorizontalSpacing(10)
        indicator_grid.setVerticalSpacing(8)

        self.session_indicator = self._make_indicator_value("—")
        self.calibration_indicator = self._make_indicator_value("Not captured")
        self.next_coord_indicator = self._make_indicator_value("143x10")
        self.slide_indicator = self._make_indicator_value("1 / 1")
        self.replicate_indicator = self._make_indicator_value(f"1 / {self.DEFAULT_PHOTOS_PER_SLIDE}")
        self.progress_indicator = self._make_indicator_value("0 complete")

        indicator_grid.addWidget(self._make_indicator_label("Session"), 0, 0)
        indicator_grid.addWidget(self.session_indicator, 0, 1)
        indicator_grid.addWidget(self._make_indicator_label("Calibration image"), 1, 0)
        indicator_grid.addWidget(self.calibration_indicator, 1, 1)
        indicator_grid.addWidget(self._make_indicator_label("Next coordinate"), 2, 0)
        indicator_grid.addWidget(self.next_coord_indicator, 2, 1)
        indicator_grid.addWidget(self._make_indicator_label("Slide number"), 3, 0)
        indicator_grid.addWidget(self.slide_indicator, 3, 1)
        indicator_grid.addWidget(self._make_indicator_label("Replicate number"), 4, 0)
        indicator_grid.addWidget(self.replicate_indicator, 4, 1)
        indicator_grid.addWidget(self._make_indicator_label("Site progress"), 5, 0)
        indicator_grid.addWidget(self.progress_indicator, 5, 1)

        self.capture_status_label = QLabel("Ready")
        self.capture_status_label.setObjectName("InfoLine")

        self.last_saved_label = QLabel("Last saved: none")
        self.last_saved_label.setObjectName("InfoLine")

        self.session_count_label = QLabel("Session images saved: 0")
        self.session_count_label.setObjectName("InfoLine")

        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setMinimumHeight(210)
        self.summary_box.setObjectName("SummaryBox")

        capture_layout.addWidget(self.capture_calibration_btn)
        capture_layout.addWidget(self.capture_btn)
        capture_layout.addLayout(indicator_grid)
        capture_layout.addWidget(self.capture_status_label)
        capture_layout.addWidget(self.last_saved_label)
        capture_layout.addWidget(self.session_count_label)
        capture_layout.addWidget(self.summary_box)

        right_panel.addWidget(capture_group)
        right_panel.addStretch(1)

        self.setStatusBar(QStatusBar())

    def _make_indicator_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("IndicatorLabel")
        return label

    def _make_indicator_value(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("IndicatorValue")
        label.setWordWrap(True)
        return label

    def apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #f3f6fb;
            }

            QWidget {
                font-family: Arial, Helvetica, sans-serif;
                font-size: 13px;
                color: #1f2937;
            }

            QLabel#TitleLabel {
                font-size: 22px;
                font-weight: 700;
                color: #0f172a;
                padding: 0 0 2px 2px;
            }

            QLabel#SubtitleLabel {
                font-size: 12px;
                color: #64748b;
                padding: 0 0 4px 2px;
            }

            QGroupBox {
                background-color: white;
                border: 1px solid #dde6f0;
                border-radius: 14px;
                margin-top: 10px;
                padding: 12px;
                font-weight: 700;
                color: #334155;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }

            QLabel#MutedValueLabel {
                color: #475569;
            }

            QLabel#SystemTimeChip {
                background-color: #eef4ff;
                color: #1d4ed8;
                border: 1px solid #c7dbff;
                border-radius: 10px;
                padding: 5px 9px;
                font-weight: 700;
            }

            QLabel#StatusChip {
                border-radius: 10px;
                padding: 5px 9px;
                font-weight: 700;
                border: 1px solid #dbe3ec;
                background-color: #f8fafc;
                color: #334155;
            }

            QLabel#InfoLine {
                color: #475569;
                font-size: 12px;
            }

            QLabel#IndicatorLabel {
                color: #64748b;
                font-size: 12px;
                font-weight: 700;
            }

            QLabel#IndicatorValue {
                background-color: #f8fafc;
                border: 1px solid #d7e2ee;
                border-radius: 10px;
                padding: 8px 10px;
                font-weight: 700;
                color: #0f172a;
            }

            QLineEdit, QSpinBox, QDoubleSpinBox, QTextEdit {
                background-color: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                padding: 6px;
            }

            QPushButton {
                background-color: #e8eef7;
                border: 1px solid #c9d6e6;
                border-radius: 10px;
                padding: 8px 10px;
                font-weight: 700;
            }

            QPushButton:hover {
                background-color: #dce7f5;
            }

            QPushButton#PrimaryButton {
                background-color: #eaf5ff;
                border: 1px solid #bddcff;
                color: #0f4c81;
            }

            QPushButton#PrimaryButton:hover {
                background-color: #d8ecff;
            }

            QPushButton#AccentButton {
                background-color: #1d4ed8;
                border: 1px solid #1e40af;
                color: white;
                font-size: 24px;
                font-weight: 800;
            }

            QPushButton#AccentButton:hover {
                background-color: #1e40af;
            }

            QPushButton#BusyAccentButton {
                background-color: #dc2626;
                border: 1px solid #991b1b;
                color: white;
                font-size: 24px;
                font-weight: 800;
            }

            QPushButton#BusyAccentButton:hover {
                background-color: #b91c1c;
            }

            QTextEdit#SummaryBox {
                background-color: #f8fafc;
                border: 1px solid #dbe3ec;
                border-radius: 10px;
            }

            QLabel#PreviewPlaceholder {
                background-color: #111827;
                color: #e5e7eb;
                font-size: 22px;
                font-weight: 600;
                border-radius: 16px;
                border: 1px solid #374151;
            }
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

    def set_capture_button_busy_style(self, busy: bool) -> None:
        self.capture_btn.setObjectName("BusyAccentButton" if busy else "AccentButton")
        self.capture_btn.style().unpolish(self.capture_btn)
        self.capture_btn.style().polish(self.capture_btn)
        self.capture_btn.update()

    # ----------------------------
    # Signals / defaults
    # ----------------------------

    def _connect_signals(self) -> None:
        self.choose_output_root_btn.clicked.connect(self.choose_output_root)
        self.capture_calibration_btn.clicked.connect(self.capture_calibration_image)
        self.capture_btn.clicked.connect(self.capture_replicate_image)

        self.prev_replicate_btn.clicked.connect(self.decrease_replicate)
        self.next_replicate_btn.clicked.connect(self.advance_replicate)
        self.next_slide_btn.clicked.connect(self.next_slide)

        self.collected_date_input.textChanged.connect(self.refresh_summary)
        self.collected_time_input.textChanged.connect(self.refresh_summary)
        self.site_code_input.textChanged.connect(self.refresh_summary)
        self.total_slides_spin.valueChanged.connect(self.handle_total_slides_changed)
        self.slide_number_spin.valueChanged.connect(self.refresh_summary)
        self.replicate_spin.valueChanged.connect(self.refresh_summary)
        self.photos_per_slide_spin.valueChanged.connect(self.handle_photos_per_slide_changed)

        self.manual_exposure_checkbox.toggled.connect(self.apply_camera_controls)
        self.shutter_spin.valueChanged.connect(self.apply_camera_controls)
        self.gain_spin.valueChanged.connect(self.apply_camera_controls)
        self.reset_camera_defaults_btn.clicked.connect(self.reset_camera_defaults)

        if self.picam2 is not None and hasattr(self.preview_widget, "done_signal"):
            self.preview_widget.done_signal.connect(self.capture_done)

    def _load_defaults(self) -> None:
        now = datetime.now()
        self.collected_date_input.setText(now.strftime("%Y%m%d"))
        self.collected_time_input.setText(now.strftime("%H%M"))
        self.slide_number_spin.setValue(1)
        self.replicate_spin.setValue(1)
        self.total_slides_spin.setValue(1)
        self.photos_per_slide_spin.setValue(self.DEFAULT_PHOTOS_PER_SLIDE)

    def update_system_time_label(self) -> None:
        self.system_time_label.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    # ----------------------------
    # Event filter
    # ----------------------------

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.modifiers() == Qt.NoModifier:
            if event.key() == Qt.Key_C and not self.capture_in_progress:
                self.capture_replicate_image()
                return True
        return super().eventFilter(obj, event)

    # ----------------------------
    # State helpers
    # ----------------------------

    def get_state(self) -> CaptureState:
        return CaptureState(
            collected_date=self.collected_date_input.text().strip(),
            collected_time=self.collected_time_input.text().strip(),
            site_code=self.site_code_input.text().strip(),
            slide_number=self.slide_number_spin.value(),
            replicate_number=self.replicate_spin.value(),
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
        if state.slide_number < 1:
            return "Slide number must be at least 1."
        if state.slide_number > self.total_slides_spin.value():
            return "Current slide cannot exceed slides in site."
        if state.replicate_number < 1 or state.replicate_number > self.max_replicates():
            return f"Replicate number must be between 1 and {self.max_replicates()}."
        return None

    def current_site_folder(self, state: CaptureState | None = None) -> Path:
        state = state or self.get_state()
        return self.output_root / state.folder_name

    def build_calibration_filename(self, state: CaptureState) -> str:
        system_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{system_stamp}_{state.site_code}_CALIBRATION.png"

    def build_replicate_filename(self, state: CaptureState, coord: tuple[int, int]) -> str:
        system_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        x, y = coord
        return (
            f"{system_stamp}_{state.site_code}"
            f"_S{state.slide_number:02d}"
            f"_R{state.replicate_number:03d}"
            f"_X{x:03d}Y{y:03d}.png"
        )

    def calibration_exists_for_site(self, state: CaptureState | None = None) -> bool:
        state = state or self.get_state()
        error = self.validate_state(state)
        if error is not None:
            return False

        folder = self.current_site_folder(state)
        return any(folder.glob("*_CALIBRATION.png"))

    def get_site_completed_capture_count(self, state: CaptureState | None = None) -> int:
        state = state or self.get_state()
        error = self.validate_state(state)
        if error is not None:
            return 0

        site_log = self.current_site_folder(state) / "replicate_log.csv"
        if not site_log.exists():
            return 0

        try:
            with open(site_log, "r", encoding="utf-8", newline="") as handle:
                row_count = sum(1 for _ in handle)
            return max(0, row_count - 1)
        except Exception:
            return 0

    def total_expected_captures(self) -> int:
        return self.total_slides_spin.value() * self.max_replicates()

    def handle_total_slides_changed(self) -> None:
        self.slide_number_spin.setMaximum(self.total_slides_spin.value())
        if self.slide_number_spin.value() > self.total_slides_spin.value():
            self.slide_number_spin.setValue(self.total_slides_spin.value())
        self.refresh_summary()

    def handle_photos_per_slide_changed(self) -> None:
        self.rebuild_grid()
        self.refresh_summary()

    # ----------------------------
    # UI actions
    # ----------------------------

    def choose_output_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder", str(self.output_root))
        if folder:
            self.output_root = Path(folder)
            self.output_root_label.setText(folder)
            self.refresh_summary()

    def advance_replicate(self) -> None:
        if self.replicate_spin.value() < self.max_replicates():
            self.replicate_spin.setValue(self.replicate_spin.value() + 1)
        else:
            self.statusBar().showMessage(f"Already at replicate {self.max_replicates()} for this slide")

    def decrease_replicate(self) -> None:
        self.replicate_spin.setValue(max(1, self.replicate_spin.value() - 1))

    def next_slide(self) -> None:
        if self.capture_in_progress:
            QMessageBox.warning(self, "Capture in progress", "Wait until the current capture is finished.")
            return

        if self.slide_number_spin.value() >= self.total_slides_spin.value():
            self.statusBar().showMessage("Already at the final slide for this site")
            return

        self.slide_number_spin.setValue(self.slide_number_spin.value() + 1)
        self.replicate_spin.setValue(1)
        self.statusBar().showMessage(f"Moved to slide {self.slide_number_spin.value()}, replicate 1")

    # ----------------------------
    # Camera controls
    # ----------------------------

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

    # ----------------------------
    # UI state
    # ----------------------------

    def set_busy_ui(self, busy: bool) -> None:
        metadata_enabled = not busy

        self.capture_calibration_btn.setEnabled(not busy)
        self.capture_btn.setEnabled(not busy)

        self.prev_replicate_btn.setEnabled(metadata_enabled)
        self.next_replicate_btn.setEnabled(metadata_enabled)
        self.next_slide_btn.setEnabled(metadata_enabled)

        self.slide_number_spin.setEnabled(metadata_enabled)
        self.replicate_spin.setEnabled(metadata_enabled)
        self.total_slides_spin.setEnabled(metadata_enabled)
        self.photos_per_slide_spin.setEnabled(metadata_enabled)

        self.set_capture_button_busy_style(busy)

    def update_capture_ui(self) -> None:
        state = self.get_state()
        calibration_exists = self.calibration_exists_for_site(state)
        completed_captures = self.get_site_completed_capture_count(state)
        total_captures = self.total_expected_captures()
        next_coord = self.coord_text(state.replicate_number)

        if self.capture_in_progress:
            self.capture_btn.setText("CAPTURING...")
            self.capture_status_label.setText(f"Capturing {next_coord}...")
        else:
            self.capture_btn.setText(f"CAPTURE {next_coord}")
            self.capture_status_label.setText("Ready")

        self.session_indicator.setText(state.folder_name if state.site_code else "Enter site metadata")
        self.calibration_indicator.setText("Captured" if calibration_exists else "Not captured")
        self.next_coord_indicator.setText(next_coord)
        self.slide_indicator.setText(f"{state.slide_number} / {self.total_slides_spin.value()}")
        self.replicate_indicator.setText(f"{state.replicate_number} / {self.max_replicates()}")
        self.progress_indicator.setText(
            f"{completed_captures} complete • {max(0, total_captures - completed_captures)} remaining"
        )

        self.session_count_label.setText(f"Session images saved: {self.session_capture_count}")

    def refresh_summary(self) -> None:
        state = self.get_state()
        error = self.validate_state(state)

        if error is not None:
            self.summary_box.setPlainText(f"Metadata check:\n{error}")
            self.update_capture_ui()
            return

        folder = self.current_site_folder(state)
        calibration_example = folder / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{state.site_code}_CALIBRATION.png"
        next_coord = self.coord_for_replicate(state.replicate_number)
        next_file = folder / self.build_replicate_filename(state, next_coord)
        completed_captures = self.get_site_completed_capture_count(state)
        total_captures = self.total_expected_captures()
        grid_rows, grid_cols = self.grid_shape_for_count(self.photos_per_slide())

        text = (
            f"Session folder:\n{folder}\n\n"
            f"Calibration image example:\n{calibration_example}\n\n"
            f"Next replicate image:\n{next_file}\n\n"
            f"Slides in site: {self.total_slides_spin.value()}\n"
            f"Current slide: {state.slide_number} of {self.total_slides_spin.value()}\n"
            f"Current replicate photo: {state.replicate_number} of {self.max_replicates()}\n"
            f"Photos per slide: {self.photos_per_slide()}\n"
            f"Next coordinate: {next_coord[0]}x{next_coord[1]}\n"
            f"Completed captures in this site: {completed_captures} of {total_captures}\n\n"
            f"Grid:\n"
            f"• Scaled snake pattern\n"
            f"• Approximate grid shape: {grid_rows} × {grid_cols}\n"
            f"• Start bounds: {self.X_START}x{self.Y_START}\n"
            f"• End bounds: {self.X_END}x{self.Y_END}\n"
            f"• Coordinates rounded to nearest mm\n"
            f"• Same area used regardless of point count\n\n"
            f"Camera mode: {'Manual' if self.manual_exposure_checkbox.isChecked() else 'Auto'}\n"
            f"Shutter: {self.shutter_spin.value()} µs\n"
            f"Gain: {self.gain_spin.value():.1f}\n\n"
            f"Capture:\n"
            f"• Capture one calibration image for the entire site\n"
            f"• Capture {self.photos_per_slide()} replicate photos per slide\n"
            f"• Press c to capture the next coordinate\n"
            f"• Images save directly into the site folder"
        )

        self.summary_box.setPlainText(text)
        self.update_capture_ui()

    # ----------------------------
    # Capture workflow
    # ----------------------------

    def begin_capture(
        self,
        full_path: Path,
        state: CaptureState,
        kind: str,
        status_text: str,
        coord: tuple[int, int] | None = None,
    ) -> None:
        self.pending_file_path = full_path
        self.pending_state = state
        self.pending_capture_kind = kind
        self.pending_capture_coord = coord
        self.capture_in_progress = True
        self.set_busy_ui(True)
        self.statusBar().showMessage(status_text)
        self.update_capture_ui()

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
            self.update_capture_ui()
            QMessageBox.critical(self, "Capture failed", f"Could not capture image:\n\n{exc}")

    def capture_calibration_image(self) -> None:
        if self.capture_in_progress:
            return

        state = self.get_state()
        error = self.validate_state(state)
        if error is not None:
            QMessageBox.warning(self, "Check metadata", error)
            return

        folder = self.current_site_folder(state)
        full_path = folder / self.build_calibration_filename(state)

        if self.calibration_exists_for_site(state):
            answer = QMessageBox.question(
                self,
                "Calibration already exists",
                "A calibration image already exists for this site.\n\nDo you want to add another calibration image?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        self.begin_capture(full_path, state, "calibration", f"Capturing calibration image {full_path.name}...")

    def capture_replicate_image(self) -> None:
        if self.capture_in_progress:
            return

        state = self.get_state()
        error = self.validate_state(state)
        if error is not None:
            QMessageBox.warning(self, "Check metadata", error)
            return

        coord = self.coord_for_replicate(state.replicate_number)
        folder = self.current_site_folder(state)
        full_path = folder / self.build_replicate_filename(state, coord)

        if full_path.exists():
            answer = QMessageBox.question(
                self,
                "Image already exists",
                f"This replicate image already exists:\n\n{full_path.name}\n\nReplace it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        self.begin_capture(
            full_path,
            state,
            "replicate",
            f"Capturing replicate {state.replicate_number} at {coord[0]}x{coord[1]}...",
            coord=coord,
        )

    def capture_done(self, job) -> None:
        try:
            if self.picam2 is not None and job is not None:
                self.picam2.wait(job)

            if self.pending_file_path is None or self.pending_state is None:
                return

            self.session_capture_count += 1
            self.last_saved_label.setText(f"Last saved: {self.pending_file_path.name}")

            if self.pending_capture_kind == "calibration":
                self.write_calibration_log(self.pending_file_path, self.pending_state)
                self.statusBar().showMessage(f"Captured calibration image: {self.pending_file_path.name}")

            elif self.pending_capture_kind == "replicate":
                coord = self.pending_capture_coord or self.coord_for_replicate(self.pending_state.replicate_number)
                self.write_replicate_log(self.pending_file_path, self.pending_state, coord)
                self.statusBar().showMessage(
                    f"Captured replicate {self.pending_state.replicate_number} at {coord[0]}x{coord[1]}"
                )

                if self.auto_increment_checkbox.isChecked():
                    self.advance_after_completed_capture(
                        self.pending_state.slide_number,
                        self.pending_state.replicate_number,
                    )

            self.capture_in_progress = False
            self.pending_file_path = None
            self.pending_state = None
            self.pending_capture_kind = "replicate"
            self.pending_capture_coord = None
            self.set_busy_ui(False)
            self.refresh_summary()

        except Exception as exc:
            QMessageBox.critical(self, "Capture completion failed", f"{exc}")
            self.capture_in_progress = False
            self.pending_file_path = None
            self.pending_state = None
            self.pending_capture_kind = "replicate"
            self.pending_capture_coord = None
            self.set_busy_ui(False)
            self.refresh_summary()

    def advance_after_completed_capture(self, slide_number: int, replicate_number: int) -> None:
        total_slides = self.total_slides_spin.value()

        if replicate_number < self.max_replicates():
            self.slide_number_spin.setValue(slide_number)
            self.replicate_spin.setValue(replicate_number + 1)
            self.statusBar().showMessage(
                f"Capture complete. Ready for slide {slide_number}, replicate {replicate_number + 1}"
            )
            return

        if slide_number < total_slides:
            self.slide_number_spin.setValue(slide_number + 1)
            self.replicate_spin.setValue(1)
            self.statusBar().showMessage(
                f"Slide {slide_number} complete. Ready for slide {slide_number + 1}, replicate 1"
            )
            return

        self.slide_number_spin.setValue(total_slides)
        self.replicate_spin.setValue(self.max_replicates())

        QMessageBox.information(
            self,
            "Site complete",
            "All slides and replicate photos for this site are complete.\n\n"
            "Please update the session metadata for the next site and capture a new calibration image."
        )

        self.statusBar().showMessage("Site complete. Ready to start next site.")

    # ----------------------------
    # Logging
    # ----------------------------

    def write_calibration_log(self, path: Path, state: CaptureState) -> None:
        log_file = path.parent / "calibration_log.csv"
        exists = log_file.exists()
        with open(log_file, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at",
                    "site_folder",
                    "site_code",
                    "filename",
                    "full_path",
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                state.folder_name,
                state.site_code,
                path.name,
                str(path),
            ])

    def write_replicate_log(self, path: Path, state: CaptureState, coord: tuple[int, int]) -> None:
        log_file = path.parent / "replicate_log.csv"
        exists = log_file.exists()
        x, y = coord
        with open(log_file, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if not exists:
                writer.writerow([
                    "saved_at",
                    "site_folder",
                    "site_code",
                    "slide_number",
                    "replicate_number",
                    "coord_x_mm",
                    "coord_y_mm",
                    "coord_text",
                    "filename",
                    "full_path",
                ])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                state.folder_name,
                state.site_code,
                state.slide_number,
                state.replicate_number,
                x,
                y,
                f"{x}x{y}",
                path.name,
                str(path),
            ])

    # ----------------------------
    # Close
    # ----------------------------

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
