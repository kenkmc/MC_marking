"""PyQt6 main window driving the MC Marking workflow."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import numpy as np
from PyQt6.QtCore import QPoint, QRect, QSize, Qt, pyqtSignal
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPen, QColor
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QComboBox,
    QInputDialog,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QCheckBox,
    QSlider,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)
from PyQt6.QtWidgets import QRubberBand

import pytesseract
from pytesseract import TesseractNotFoundError

from mc_marking.models.answer_sheet import AnswerKey, CellResult, PageAnswer, PageResult, TableExtraction
from mc_marking.services.answer_key_service import normalize_answer_key
from mc_marking.services.image_loader import LoadedPage, load_pages
from mc_marking.services.marking_service import evaluate_page
from mc_marking.services.ocr_service import recognise_table_cells
from mc_marking.services.page_processor import build_page_result
from mc_marking.services.table_detection import detect_table, detect_tables
from mc_marking.services.table_parser import enumerate_question_marks
from mc_marking.utils.image_utils import BoundingBox
from mc_marking.utils.settings import AppSettings, load_settings, save_settings

SUPPORTED_FILTER = "PDF or Images (*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff)"


class ImageCanvas(QLabel):
    """Image display widget that supports manual rectangular selection."""

    selection_changed = pyqtSignal(object)

    def __init__(self) -> None:
        super().__init__()
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setScaledContents(False)
        self._rubber_band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self._origin: Optional[QPoint] = None
        self._selection: Optional[QRect] = None
        self._image: Optional[np.ndarray] = None
        self._original_pixmap: Optional[QPixmap] = None
        self._zoom_factor: float = 1.0
        self._overlay_boxes: Dict[str, List[BoundingBox]] = {}
        self._overlay_visibility: Dict[str, bool] = {}
        self._overlay_colors: Dict[str, QColor] = {
            "questions": QColor(255, 215, 0),  # gold
            "choices": QColor(30, 144, 255),   # dodger blue
            "answers": QColor(220, 20, 60),    # crimson
            "ocr_regions": QColor(60, 179, 113),  # medium sea green
            "omr_regions": QColor(138, 43, 226),  # blue violet
        }

    def set_image(self, image: np.ndarray) -> None:
        self._image = image
        self._selection = None
        self._rubber_band.hide()
        qimage = _np_to_qimage(image)
        pixmap = QPixmap.fromImage(qimage)
        self._original_pixmap = pixmap
        self._zoom_factor = 1.0
        self._update_display_pixmap()
        self.clear_overlays()

    def clear(self) -> None:
        self._image = None
        self._selection = None
        self._rubber_band.hide()
        self.setPixmap(QPixmap())
        self._original_pixmap = None
        self._zoom_factor = 1.0
        self.clear_overlays()

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if self._image is None:
            return
        self._origin = event.position().toPoint()
        self._rubber_band.setGeometry(QRect(self._origin, QSize()))
        self._rubber_band.show()

    def mouseMoveEvent(self, event) -> None:  # type: ignore[override]
        if self._origin is None or self._image is None:
            return
        current = event.position().toPoint()
        rect = QRect(self._origin, current).normalized()
        self._rubber_band.setGeometry(rect)

    def mouseReleaseEvent(self, event) -> None:  # type: ignore[override]
        if self._origin is None or self._image is None:
            return
        current = event.position().toPoint()
        self._selection = QRect(self._origin, current).normalized()
        self.selection_changed.emit(self.get_selection_box())
        self._origin = None

    def get_selection_box(self) -> Optional[BoundingBox]:
        if self._image is None or self._selection is None or self._selection.isNull():
            return None
        display_rect = self._current_pixmap_rect()
        if display_rect is None or display_rect.width() <= 0 or display_rect.height() <= 0:
            return None
        rect = self._selection.intersected(display_rect)
        if rect.isNull():
            return None
        image_height, image_width = self._image.shape[:2]
        scale_x = image_width / display_rect.width()
        scale_y = image_height / display_rect.height()
        x = int((rect.x() - display_rect.x()) * scale_x)
        y = int((rect.y() - display_rect.y()) * scale_y)
        width = int(rect.width() * scale_x)
        height = int(rect.height() * scale_y)
        x = max(0, min(x, image_width - 1))
        y = max(0, min(y, image_height - 1))
        if x + width > image_width:
            width = image_width - x
        if y + height > image_height:
            height = image_height - y
        width = max(1, width)
        height = max(1, height)
        return BoundingBox(x=x, y=y, width=width, height=height)

    def original_pixmap(self) -> Optional[QPixmap]:
        return self._original_pixmap

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._update_display_pixmap()

    def set_zoom(self, factor: float) -> None:
        self._zoom_factor = max(0.1, factor)
        self._selection = None
        self._rubber_band.hide()
        self._update_display_pixmap()

    def update_overlays(self, overlays: Dict[str, List[BoundingBox]]) -> None:
        self._overlay_boxes = {category: list(boxes) for category, boxes in overlays.items()}
        for category in overlays:
            self._overlay_visibility.setdefault(category, True)
        self.update()

    def clear_overlays(self) -> None:
        self._overlay_boxes = {}
        self.update()

    def set_overlay_visibility(self, category: str, visible: bool) -> None:
        self._overlay_visibility[category] = visible
        self.update()

    def overlay_visible(self, category: str) -> bool:
        return self._overlay_visibility.get(category, True)

    def paintEvent(self, event) -> None:  # type: ignore[override]
        super().paintEvent(event)
        if not self._overlay_boxes or self.pixmap() is None or self.pixmap().isNull():
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        for category, boxes in self._overlay_boxes.items():
            if not self._overlay_visibility.get(category, True):
                continue
            color = self._overlay_colors.get(category, QColor(255, 0, 0))
            pen = QPen(color)
            pen.setWidth(2)
            painter.setPen(pen)
            for box in boxes:
                rect = self._map_box_to_widget(box)
                if rect is not None:
                    painter.drawRect(rect)
        painter.end()

    def _map_box_to_widget(self, box: BoundingBox) -> Optional[QRect]:
        if self._image is None or box.width <= 0 or box.height <= 0:
            return None
        display_rect = self._current_pixmap_rect()
        if display_rect is None or display_rect.width() <= 0 or display_rect.height() <= 0:
            return None
        image_height, image_width = self._image.shape[:2]
        scale_x = display_rect.width() / image_width
        scale_y = display_rect.height() / image_height
        x = display_rect.x() + int(box.x * scale_x)
        y = display_rect.y() + int(box.y * scale_y)
        width = max(1, int(box.width * scale_x))
        height = max(1, int(box.height * scale_y))
        return QRect(x, y, width, height)

    def _update_display_pixmap(self) -> None:
        if self._original_pixmap is None or self._original_pixmap.isNull():
            self.setPixmap(QPixmap())
            return
        base_width = self._original_pixmap.width()
        base_height = self._original_pixmap.height()
        if base_width <= 0 or base_height <= 0:
            return
        label_width = self.width()
        label_height = self.height()
        if label_width <= 0 or label_height <= 0:
            super().setPixmap(self._original_pixmap)
            self.update()
            return
        label_width = max(1, label_width)
        label_height = max(1, label_height)
        fit_scale = min(label_width / base_width, label_height / base_height)
        if fit_scale <= 0:
            fit_scale = 1.0
        fit_scale = min(fit_scale, 1.0)
        scale = fit_scale * self._zoom_factor
        target_width = max(1, int(base_width * scale))
        target_height = max(1, int(base_height * scale))
        scaled = self._original_pixmap.scaled(
            target_width,
            target_height,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        super().setPixmap(scaled)
        self.update()

    def _current_pixmap_rect(self) -> Optional[QRect]:
        pixmap = self.pixmap()
        if pixmap is None or pixmap.isNull():
            return None
        width = pixmap.width()
        height = pixmap.height()
        if width <= 0 or height <= 0:
            return None
        label_width = self.width()
        label_height = self.height()
        x_offset = (label_width - width) // 2
        y_offset = (label_height - height) // 2
        return QRect(x_offset, y_offset, width, height)


class PreviewCanvasContainer(QFrame):
    """Fixed-size frame that hosts the preview canvas and zoom controls."""

    zoom_in_requested = pyqtSignal()
    zoom_out_requested = pyqtSignal()

    def __init__(self, canvas: ImageCanvas, *, fixed_size: QSize = QSize(640, 720)) -> None:
        super().__init__()
        self._canvas = canvas
        self._fixed_size = fixed_size
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.setFixedSize(self._fixed_size)

        self._canvas.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._canvas.setFixedSize(self._fixed_size)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._canvas)
        self.setLayout(layout)

        self._overlay = QFrame(self)
        self._overlay.setFrameShape(QFrame.Shape.NoFrame)
        self._overlay.setStyleSheet("background-color: rgba(0, 0, 0, 110); border-radius: 6px;")
        self._overlay.raise_()

        overlay_layout = QHBoxLayout()
        overlay_layout.setContentsMargins(10, 6, 10, 6)
        overlay_layout.setSpacing(6)
        self._overlay.setLayout(overlay_layout)

        self.zoom_out_button = QToolButton(self._overlay)
        self.zoom_out_button.setText("-")
        self.zoom_out_button.setToolTip("Zoom out")
        self.zoom_out_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.zoom_out_button.setStyleSheet(
            "QToolButton { color: white; background-color: rgba(255, 255, 255, 70);"
            " border: none; padding: 4px 10px; border-radius: 4px; }"
            "QToolButton:hover { background-color: rgba(255, 255, 255, 110); color: black; }"
        )
        self.zoom_out_button.clicked.connect(self.zoom_out_requested.emit)

        self.zoom_in_button = QToolButton(self._overlay)
        self.zoom_in_button.setText("+")
        self.zoom_in_button.setToolTip("Zoom in")
        self.zoom_in_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.zoom_in_button.setStyleSheet(
            "QToolButton { color: white; background-color: rgba(255, 255, 255, 70);"
            " border: none; padding: 4px 10px; border-radius: 4px; }"
            "QToolButton:hover { background-color: rgba(255, 255, 255, 110); color: black; }"
        )
        self.zoom_in_button.clicked.connect(self.zoom_in_requested.emit)

        overlay_layout.addWidget(self.zoom_out_button)
        overlay_layout.addWidget(self.zoom_in_button)

        self._position_overlay()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._position_overlay()

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        self._position_overlay()

    def _position_overlay(self) -> None:
        self._overlay.adjustSize()
        margin = 12
        x_pos = max(margin, self.width() - self._overlay.width() - margin)
        y_pos = margin
        self._overlay.move(x_pos, y_pos)


class MainWindow(QMainWindow):
    """Top-level window orchestrating answer key capture and marking."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MC Marking Tool")
        self.setGeometry(100, 100, 800, 600)

        # Main layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # File import button
        self.import_button = QPushButton("Import PDF")
        self.import_button.clicked.connect(self.import_pdf)
        self.layout.addWidget(self.import_button)

        # Answer key checkbox
        self.answer_key_checkbox = QCheckBox("Is the first page the answer key?")
        self.layout.addWidget(self.answer_key_checkbox)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.settings: AppSettings = load_settings()
        self._apply_runtime_paths()
        self.baseline_density: Optional[float] = self.settings.baseline_ink_density
        self.row_labels_override: List[str] = list(self.settings.row_labels)
        self.question_column_template: Optional[Dict[str, float]] = (
            dict(self.settings.question_column_template) if self.settings.question_column_template else None
        )
        self.choice_row_templates: List[Dict[str, float]] = [dict(item) for item in self.settings.choice_row_templates]
        self.layout_capture_mode: Optional[str] = None
        self._pending_row_capture_labels: List[str] = []
        self._captured_row_boxes: List[Dict[str, float]] = []
        self._captured_question_box: Optional[Dict[str, float]] = None
        self._reference_table_box: Optional[BoundingBox] = None
        self._calibration_page: Optional[LoadedPage] = None
        self.last_selection: Optional[BoundingBox] = None
        self._capture_labels_used: List[str] = []
        self.layout_templates: Dict[str, Dict[str, Any]] = {}
        self.active_layout_name: Optional[str] = None
        self._updating_layout_selector = False
        self.ocr_region_templates: List[Dict[str, float]] = [dict(box) for box in self.settings.ocr_regions]
        self.omr_region_templates: List[Dict[str, float]] = [dict(box) for box in self.settings.omr_regions]
        self.region_capture_mode: Optional[str] = None

        self.answer_key: Optional[AnswerKey] = None
        self.answer_extractions: List[TableExtraction] = []
        self.answer_page: Optional[LoadedPage] = None
        self.answer_cells: List[List[CellResult]] = []
        self.page_results: List[PageResult] = []

        self.canvas = ImageCanvas()
        self.canvas.selection_changed.connect(self._on_selection_changed)
        self.preview_container = PreviewCanvasContainer(self.canvas)
        self.preview_container.zoom_in_requested.connect(lambda: self._step_zoom(10))
        self.preview_container.zoom_out_requested.connect(lambda: self._step_zoom(-10))

        self.answer_table = QTableWidget(0, 2)
        self.answer_table.setHorizontalHeaderLabels(["Question", "Answer"])
        self.answer_table.horizontalHeader().setStretchLastSection(True)

        self.answer_log = QTextEdit()
        self.answer_log.setReadOnly(True)

        self.results_table = QTableWidget(0, 5)
        self.results_table.setHorizontalHeaderLabels(["Source", "Correct", "Incorrect", "Unanswered", "Total"])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.results_table.itemSelectionChanged.connect(self._on_page_result_selected)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.results_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.results_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        self.student_answers_table = QTableWidget(0, 4)
        self.student_answers_table.setHorizontalHeaderLabels(["Question", "Detected", "Confidence", "Status"])
        self.student_answers_table.horizontalHeader().setStretchLastSection(True)
        self.student_answers_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.student_answers_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.student_answers_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(10, 300)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(self._on_zoom_changed)
        self.reset_zoom_button = QPushButton("Reset Zoom")
        self.reset_zoom_button.clicked.connect(self._reset_zoom)
        zoom_row = QHBoxLayout()
        zoom_row.addWidget(QLabel("Zoom"))
        zoom_row.addWidget(self.zoom_slider)
        zoom_row.addWidget(self.reset_zoom_button)

        self.show_question_boxes = QCheckBox("Questions")
        self.show_question_boxes.setChecked(True)
        self.show_choice_boxes = QCheckBox("Choices")
        self.show_choice_boxes.setChecked(True)
        self.show_answer_boxes = QCheckBox("Answers")
        self.show_answer_boxes.setChecked(True)
        self.show_ocr_regions = QCheckBox("OCR Zones")
        self.show_ocr_regions.setChecked(True)
        self.show_omr_regions = QCheckBox("OMR Zones")
        self.show_omr_regions.setChecked(True)
        overlays_row = QHBoxLayout()
        overlays_row.addWidget(QLabel("Overlays"))
        overlays_row.addWidget(self.show_question_boxes)
        overlays_row.addWidget(self.show_choice_boxes)
        overlays_row.addWidget(self.show_answer_boxes)
        overlays_row.addWidget(self.show_ocr_regions)
        overlays_row.addWidget(self.show_omr_regions)
        overlays_row.addStretch(1)

        self._overlay_preferences: Dict[str, bool] = {
            "questions": True,
            "choices": True,
            "answers": True,
            "ocr_regions": True,
            "omr_regions": True,
        }
        self._overlay_checkboxes: Dict[str, QCheckBox] = {
            "questions": self.show_question_boxes,
            "choices": self.show_choice_boxes,
            "answers": self.show_answer_boxes,
            "ocr_regions": self.show_ocr_regions,
            "omr_regions": self.show_omr_regions,
        }
        for category, checkbox in self._overlay_checkboxes.items():
            checkbox.setEnabled(False)
            checkbox.toggled.connect(lambda checked, cat=category: self._on_overlay_toggled(cat, checked))

        self.calibrate_button = QPushButton("Calibrate Blank Sheet")
        self.calibrate_button.clicked.connect(self._calibrate_blank_sheet)

        self.configure_poppler_button = QPushButton("Set Poppler Path")
        self.configure_poppler_button.clicked.connect(self._configure_poppler_path)

        self.configure_tesseract_button = QPushButton("Set Tesseract Path")
        self.configure_tesseract_button.clicked.connect(self._configure_tesseract_path)

        self.load_answer_button = QPushButton("Load Answer Key")
        self.load_answer_button.clicked.connect(self._load_answer_key)

        self.use_selection_button = QPushButton("Use Selection")
        self.use_selection_button.setEnabled(False)
        self.use_selection_button.clicked.connect(self._apply_manual_selection)

        self.process_button = QPushButton("Process Answer Sheets")
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self._process_answer_sheets)

        self.add_ocr_region_button = QPushButton("Add OCR Region")
        self.add_ocr_region_button.clicked.connect(lambda: self._start_region_capture("ocr"))
        self.add_omr_region_button = QPushButton("Add OMR Region")
        self.add_omr_region_button.clicked.connect(lambda: self._start_region_capture("omr"))
        self.clear_region_button = QPushButton("Clear Region Overrides")
        self.clear_region_button.clicked.connect(self._clear_region_templates)

        self.layout_selector = QComboBox()
        self.layout_selector.currentIndexChanged.connect(self._on_layout_selected)
        self.delete_layout_button = QPushButton("Delete Layout")
        self.delete_layout_button.setEnabled(False)
        self.delete_layout_button.clicked.connect(self._delete_layout)

        self.configure_layout_button = QPushButton("Configure Layout")
        self.configure_layout_button.clicked.connect(self._configure_row_labels)

        self.export_button = QPushButton("Export Results")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self._export_results)

        self.mark_button = QPushButton("Mark Text Fields")
        self.mark_button.clicked.connect(self.start_marking)
        self.layout.addWidget(self.mark_button)

        self._load_layout_templates()

        layout_selector_row = QHBoxLayout()
        layout_selector_row.addWidget(QLabel("Layout"))
        layout_selector_row.addWidget(self.layout_selector)
        layout_selector_widget = QWidget()
        layout_selector_widget.setLayout(layout_selector_row)

        left_panel = QVBoxLayout()
        left_panel.addWidget(self.configure_poppler_button)
        left_panel.addWidget(self.configure_tesseract_button)
        left_panel.addWidget(self.load_answer_button)
        left_panel.addWidget(self.use_selection_button)
        left_panel.addWidget(self.process_button)
        left_panel.addWidget(layout_selector_widget)
        left_panel.addWidget(self.configure_layout_button)
        left_panel.addWidget(self.delete_layout_button)
        left_panel.addWidget(self.export_button)
        left_panel.addWidget(self.calibrate_button)
        left_panel.addWidget(self.add_ocr_region_button)
        left_panel.addWidget(self.add_omr_region_button)
        left_panel.addWidget(self.clear_region_button)

        left_panel.addWidget(_build_group_box("Answer Key", self.answer_table))
        left_panel.addWidget(_build_group_box("Recognition Log", self.answer_log))
        left_panel_widget = QWidget()
        left_panel_widget.setLayout(left_panel)

        central_layout = QHBoxLayout()
        central_layout.addWidget(left_panel_widget, 1)

        preview_panel = QVBoxLayout()
        preview_panel.setContentsMargins(0, 0, 0, 0)
        preview_group = _build_group_box("Preview", self.preview_container)
        preview_group.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        preview_panel.addWidget(preview_group, alignment=Qt.AlignmentFlag.AlignTop)
        preview_panel.addLayout(zoom_row)
        preview_panel.addLayout(overlays_row)
        preview_widget = QWidget()
        preview_widget.setLayout(preview_panel)
        preview_widget.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
        central_layout.addWidget(preview_widget)

        right_panel = QVBoxLayout()
        right_panel.addWidget(_build_group_box("Page Results", self.results_table))
        right_panel.addWidget(_build_group_box("Student Answers", self.student_answers_table))
        right_panel_widget = QWidget()
        right_panel_widget.setLayout(right_panel)
        central_layout.addWidget(right_panel_widget, 1)

        central_widget = QWidget()
        central_widget.setLayout(central_layout)
        self.setCentralWidget(central_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self._clear_preview_overlays()
        self._update_region_buttons_state()

    def enable_marking(self):
        self.mark_button = QPushButton("Mark Text Fields")
        self.mark_button.clicked.connect(self.start_marking)
        self.layout.addWidget(self.mark_button)

        self.canvas = ImageCanvas()
        self.layout.addWidget(self.canvas)

    def start_marking(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Image for Marking", "", SUPPORTED_FILTER)
        if file_path:
            self.status_bar.showMessage(f"Marking: {file_path}")
            image = self.load_image(file_path)
            self.canvas.set_image(image)

    def load_image(self, file_path: str):
        # Placeholder for image loading logic
        # Replace with actual image loading code
        return np.zeros((800, 600, 3), dtype=np.uint8)

    def _on_overlay_toggled(self, category: str, checked: bool) -> None:
        self._overlay_preferences[category] = checked
        self.canvas.set_overlay_visibility(category, checked)

    def _update_overlay_controls(self, overlays: Dict[str, List[BoundingBox]]) -> None:
        for category, checkbox in self._overlay_checkboxes.items():
            has_boxes = bool(overlays.get(category))
            desired = self._overlay_preferences.get(category, True)
            checkbox.blockSignals(True)
            if has_boxes:
                checkbox.setChecked(desired)
            else:
                checkbox.setChecked(False)
            checkbox.setEnabled(has_boxes)
            checkbox.blockSignals(False)
            if has_boxes:
                self.canvas.set_overlay_visibility(category, checkbox.isChecked())
            else:
                self.canvas.set_overlay_visibility(category, False)

    def _clear_preview_overlays(self) -> None:
        self.canvas.clear_overlays()
        empty_overlays: Dict[str, List[BoundingBox]] = {
            "questions": [],
            "choices": [],
            "answers": [],
            "ocr_regions": [],
            "omr_regions": [],
        }
        self._update_overlay_controls(empty_overlays)

    def _build_overlay_map(
        self,
        extractions: Sequence[TableExtraction],
        tables_cells: Sequence[Sequence[CellResult]],
    ) -> Dict[str, List[BoundingBox]]:
        overlay_map: Dict[str, List[BoundingBox]] = {
            "questions": [],
            "choices": [],
            "answers": [],
            "ocr_regions": [],
            "omr_regions": [],
        }
        for extraction, cells in zip(extractions, tables_cells):
            for cell in cells:
                shrunk_box = self._shrink_overlay_box(cell.bounding_box)
                if cell.row == 0 and cell.column > 0:
                    overlay_map["questions"].append(shrunk_box)
                elif cell.row > 0:
                    overlay_map["choices"].append(shrunk_box)
        for _, question in enumerate_question_marks(
            extractions,
            tables_cells,
            baseline_density=self.baseline_density,
            row_labels_override=self.row_labels_override,
        ):
            for _, cell in question.marked:
                overlay_map["answers"].append(self._shrink_overlay_box(cell.bounding_box, scale=0.7))
        for extraction in extractions:
            if extraction.bounding_box is None:
                continue
            overlay_map["ocr_regions"].extend(
                self._convert_templates_to_boxes(self.ocr_region_templates, extraction.bounding_box)
            )
            overlay_map["omr_regions"].extend(
                self._convert_templates_to_boxes(self.omr_region_templates, extraction.bounding_box)
            )
        return overlay_map

    @staticmethod
    def _shrink_overlay_box(box: BoundingBox, *, scale: float = 0.6) -> BoundingBox:
        scale = max(0.05, min(scale, 1.0))
        new_width = max(1, int(box.width * scale))
        new_height = max(1, int(box.height * scale))
        offset_x = (box.width - new_width) // 2
        offset_y = (box.height - new_height) // 2
        return BoundingBox(
            x=box.x + offset_x,
            y=box.y + offset_y,
            width=new_width,
            height=new_height,
        )

    @staticmethod
    def _convert_templates_to_boxes(
        templates: Sequence[Dict[str, float]],
        reference: Optional[BoundingBox],
    ) -> List[BoundingBox]:
        boxes: List[BoundingBox] = []
        if reference is None:
            return boxes
        for template in templates:
            try:
                boxes.append(MainWindow._denormalize_relative_box(template, reference))
            except (KeyError, TypeError):
                continue
        return boxes

    def _apply_preview_overlays(
        self,
        extractions: Sequence[TableExtraction],
        tables_cells: Sequence[Sequence[CellResult]],
    ) -> None:
        if not extractions or not tables_cells:
            self._clear_preview_overlays()
            return
        overlay_map = self._build_overlay_map(extractions, tables_cells)
        self.canvas.update_overlays(overlay_map)
        self._update_overlay_controls(overlay_map)

    def _start_region_capture(self, mode: str) -> None:
        if self.layout_capture_mode is not None:
            QMessageBox.warning(
                self,
                "Region Capture",
                "Finish or cancel the current layout capture before adding region overrides.",
            )
            return
        if not self.answer_extractions or all(extraction.bounding_box is None for extraction in self.answer_extractions):
            QMessageBox.information(
                self,
                "Region Capture",
                "Load an answer key so the table bounds are known before defining OCR/OMR regions.",
            )
            return
        label = "OCR" if mode == "ocr" else "OMR"
        self.region_capture_mode = mode
        self.status_bar.showMessage(
            f"Draw a rectangle for the {label} region, then release to save it.",
            8000,
        )
        QMessageBox.information(
            self,
            "Region Capture",
            f"Use the preview to draw a rectangle covering the {label} zone, then release to confirm.",
        )

    def _handle_region_selection(self, selection: BoundingBox) -> bool:
        if self.region_capture_mode is None:
            return False
        if not self.answer_extractions:
            self.region_capture_mode = None
            return True

        chosen_relative: Optional[Dict[str, float]] = None
        for extraction in self.answer_extractions:
            if extraction.bounding_box is None:
                continue
            relative = self._compute_relative_box(selection, extraction.bounding_box)
            if relative is None or relative.get("width", 0.0) <= 0.0 or relative.get("height", 0.0) <= 0.0:
                continue
            chosen_relative = relative
            break

        if chosen_relative is None:
            QMessageBox.warning(
                self,
                "Region Capture",
                "Selection must overlap a detected table. Please try again.",
            )
            return True

        if self.region_capture_mode == "ocr":
            self.ocr_region_templates.append(chosen_relative)
            region_label = "OCR"
        else:
            self.omr_region_templates.append(chosen_relative)
            region_label = "OMR"

        self.region_capture_mode = None
        self._update_region_templates_in_settings()
        self._refresh_region_overlays()
        self._update_region_buttons_state()
        self.status_bar.showMessage(f"{region_label} region added.", 5000)
        return True

    def _resolve_region_assignments(
        self,
        extraction: TableExtraction,
    ) -> Optional[Dict[str, List[BoundingBox]]]:
        if extraction.bounding_box is None:
            return None
        ocr_boxes = self._convert_templates_to_boxes(self.ocr_region_templates, extraction.bounding_box)
        omr_boxes = self._convert_templates_to_boxes(self.omr_region_templates, extraction.bounding_box)
        if not ocr_boxes and not omr_boxes:
            return None
        return {
            "ocr": ocr_boxes,
            "omr": omr_boxes,
        }

    def _update_region_templates_in_settings(self) -> None:
        self.settings.ocr_regions = [dict(box) for box in self.ocr_region_templates]
        self.settings.omr_regions = [dict(box) for box in self.omr_region_templates]
        if self.active_layout_name and self.active_layout_name in self.layout_templates:
            layout = self.layout_templates[self.active_layout_name]
            layout["ocr_regions"] = [dict(box) for box in self.ocr_region_templates]
            layout["omr_regions"] = [dict(box) for box in self.omr_region_templates]
            self.settings.layout_templates = self._serialize_layout_templates()

        save_settings(self.settings)

    def _refresh_region_overlays(self) -> None:
        if self.answer_extractions and self.answer_cells:
            self._apply_preview_overlays(self.answer_extractions, self.answer_cells)

    def _update_region_buttons_state(self) -> None:
        has_regions = bool(self.ocr_region_templates or self.omr_region_templates)
        self.clear_region_button.setEnabled(has_regions)

    def _clear_region_templates(self) -> None:
        if not self.ocr_region_templates and not self.omr_region_templates:
            self.status_bar.showMessage("No region overrides to clear.", 4000)
            return
        confirm = QMessageBox.question(
            self,
            "Clear Regions",
            "Remove all OCR/OMR region overrides?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.ocr_region_templates = []
        self.omr_region_templates = []
        self.region_capture_mode = None
        self._update_region_templates_in_settings()
        self._refresh_region_overlays()
        self._update_region_buttons_state()
        self.status_bar.showMessage("Region overrides cleared.", 5000)

    def _load_answer_key(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Select answer key", str(Path.home()), SUPPORTED_FILTER)
        if not file_name:
            return
        try:
            pages = load_pages([Path(file_name)], poppler_path=self.settings.poppler_path)
        except Exception as exc:  # pragma: no cover - GUI feedback
            QMessageBox.critical(self, "Load Error", f"Failed to load file: {exc}")
            return
        if not pages:
            QMessageBox.warning(self, "No Pages", "No pages were extracted from the selected file.")
            return

        page = pages[0]
        self.answer_page = page
        self.canvas.set_image(page.image)
        self._reset_zoom()
        self._clear_preview_overlays()
        self.status_bar.showMessage("Detecting tables in answer key...")
        extractions = detect_tables(page.image, page.page_index, str(page.source))
        if not extractions:
            self.answer_log.setPlainText(
                "Automatic detection failed. Draw a rectangle around the tables and choose 'Use Selection'."
            )
            self.use_selection_button.setEnabled(True)
            self.process_button.setEnabled(False)
            self.status_bar.showMessage("Detection failed; awaiting manual selection.")
            return
        self._finalize_answer_key(page, extractions)

    def _apply_manual_selection(self) -> None:
        if self.answer_page is None or self.answer_page.image is None:
            return
        selection = self.canvas.get_selection_box()
        if selection is None:
            QMessageBox.information(self, "No Selection", "Draw a rectangle over the table first.")
            return
        self.status_bar.showMessage("Running detection on selected area...")
        self._clear_preview_overlays()
        extractions = detect_tables(
            self.answer_page.image,
            self.answer_page.page_index,
            str(self.answer_page.source),
            roi=selection,
        )
        if not extractions:
            self.status_bar.showMessage("Detection still failed. Adjust the selection and try again.")
            return
        self._finalize_answer_key(self.answer_page, extractions)

    def _finalize_answer_key(self, page: LoadedPage, extractions: List[TableExtraction]) -> None:
        tables_cells: List[List[CellResult]] = []
        try:
            for extraction in extractions:
                region_config = self._resolve_region_assignments(extraction)
                ocr_results = recognise_table_cells(page.image, extraction, region_config=region_config)
                extraction.cells = ocr_results
                tables_cells.append(ocr_results)
        except TesseractNotFoundError:
            QMessageBox.critical(
                self,
                "Tesseract Not Found",
                "Tesseract OCR executable could not be located. Install Tesseract or choose it via 'Set Tesseract Path'.",
            )
            self.status_bar.showMessage("Set the Tesseract path before proceeding.", 5000)
            return
        answer_key = normalize_answer_key(
            extractions,
            tables_cells,
            baseline_density=self.baseline_density,
            row_labels_override=self.row_labels_override,
        )
        if not answer_key.rows:
            self.answer_log.setPlainText(
                "No valid rows were recognized. Ensure the table has question numbers in the first column."
            )
            self.status_bar.showMessage("Recognition incomplete.")
            return
        self.answer_key = answer_key
        self.answer_extractions = list(extractions)
        self.answer_cells = tables_cells
        self.page_results = []
        self.results_table.setRowCount(0)
        self.student_answers_table.setRowCount(0)
        self.use_selection_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.export_button.setEnabled(False)
        self._populate_answer_table(answer_key)
        log_lines = [f"Q{question}: {answer}" for question, answer in sorted(answer_key.rows.items())]
        self.answer_log.setPlainText("\n".join(log_lines))
        self.status_bar.showMessage(f"Captured answer key from {len(extractions)} table(s).", 5000)
        self._apply_preview_overlays(extractions, tables_cells)

    def _process_answer_sheets(self) -> None:
        if self.answer_key is None:
            QMessageBox.information(self, "Answer Key Needed", "Load an answer key before processing sheets.")
            return
        files, _ = QFileDialog.getOpenFileNames(self, "Select answer sheets", str(Path.home()), SUPPORTED_FILTER)
        if not files:
            return
        try:
            pages = load_pages((Path(file) for file in files), poppler_path=self.settings.poppler_path)
        except Exception as exc:  # pragma: no cover - GUI feedback
            QMessageBox.critical(self, "Load Error", f"Failed to load pages: {exc}")
            return
        results: List[PageResult] = []
        failures: List[str] = []
        for page in pages:
            extractions = detect_tables(page.image, page.page_index, str(page.source))
            if not extractions and self.answer_extractions:
                extractions = []
                for reference in self.answer_extractions:
                    if reference.bounding_box is None:
                        continue
                    fallback = detect_table(
                        page.image,
                        page.page_index,
                        str(page.source),
                        roi=reference.bounding_box,
                    )
                    if fallback is not None and fallback.cells:
                        extractions.append(fallback)
            if not extractions:
                failures.append(f"{page.source.name} (page {page.page_index + 1})")
                continue
            tables_cells: List[List[CellResult]] = []
            try:
                for extraction in extractions:
                    region_config = self._resolve_region_assignments(extraction)
                    ocr_cells = recognise_table_cells(page.image, extraction, region_config=region_config)
                    extraction.cells = ocr_cells
                    tables_cells.append(ocr_cells)
            except TesseractNotFoundError:
                QMessageBox.critical(
                    self,
                    "Tesseract Not Found",
                    "Processing stopped because Tesseract is unavailable. Install it or set the executable path.",
                )
                self.status_bar.showMessage("Set the Tesseract path before processing.", 5000)
                return
            page_result = build_page_result(
                extractions,
                tables_cells,
                baseline_density=self.baseline_density,
                row_labels_override=self.row_labels_override,
            )
            evaluated = evaluate_page(page_result, self.answer_key)
            results.append(evaluated)
        self.page_results = results
        self._populate_results_table()
        if self.page_results:
            self.results_table.selectRow(0)
        else:
            self.student_answers_table.setRowCount(0)
        self.export_button.setEnabled(bool(self.page_results))
        if failures:
            self.answer_log.append("\nUnprocessed pages:\n" + "\n".join(failures))
        self.status_bar.showMessage(f"Processed {len(results)} pages.", 5000)

    def _populate_answer_table(self, answer_key: AnswerKey) -> None:
        rows = sorted(answer_key.rows.items())
        self.answer_table.setRowCount(len(rows))
        for row_idx, (question, answer) in enumerate(rows):
            self.answer_table.setItem(row_idx, 0, QTableWidgetItem(str(question)))
            self.answer_table.setItem(row_idx, 1, QTableWidgetItem(answer))

    def _on_zoom_changed(self, value: int) -> None:
        self.canvas.set_zoom(value / 100.0)

    def _step_zoom(self, delta: int) -> None:
        new_value = self.zoom_slider.value() + delta
        new_value = max(self.zoom_slider.minimum(), min(self.zoom_slider.maximum(), new_value))
        if new_value != self.zoom_slider.value():
            self.zoom_slider.setValue(new_value)

    def _reset_zoom(self) -> None:
        self.zoom_slider.blockSignals(True)
        self.zoom_slider.setValue(100)
        self.zoom_slider.blockSignals(False)
        self.canvas.set_zoom(1.0)

    def _populate_results_table(self) -> None:
        self.results_table.blockSignals(True)
        self.results_table.setRowCount(len(self.page_results))
        for row_idx, result in enumerate(self.page_results):
            self.results_table.setItem(row_idx, 0, QTableWidgetItem(f"{result.source_path.name} (p{result.page_index + 1})"))
            self.results_table.setItem(row_idx, 1, QTableWidgetItem(str(result.correct_count)))
            self.results_table.setItem(row_idx, 2, QTableWidgetItem(str(result.incorrect_count)))
            self.results_table.setItem(row_idx, 3, QTableWidgetItem(str(result.unanswered_count)))
            self.results_table.setItem(row_idx, 4, QTableWidgetItem(str(result.total_questions)))
        self.results_table.blockSignals(False)

    def _on_page_result_selected(self) -> None:
        selection_model = self.results_table.selectionModel()
        if selection_model is None:
            return
        selected_rows = selection_model.selectedRows()
        if not selected_rows:
            self.student_answers_table.setRowCount(0)
            return
        row = selected_rows[0].row()
        if 0 <= row < len(self.page_results):
            self._populate_student_answers(self.page_results[row])

    def _populate_student_answers(self, page_result: PageResult) -> None:
        answers = page_result.answers
        self.student_answers_table.setRowCount(len(answers))
        for row_idx, answer in enumerate(answers):
            self.student_answers_table.setItem(row_idx, 0, QTableWidgetItem(str(answer.question)))
            self.student_answers_table.setItem(row_idx, 1, QTableWidgetItem(answer.extracted))
            self.student_answers_table.setItem(row_idx, 2, QTableWidgetItem(f"{answer.confidence:.2f}"))
            expected = self.answer_key.answer_for(answer.question) if self.answer_key else ""
            if answer.is_correct is True:
                status = "Correct"
            elif answer.is_correct is False:
                status = f"Incorrect (expected {expected})" if expected else "Incorrect"
            else:
                status = "Unanswered" if not answer.extracted else "Review"
            self.student_answers_table.setItem(row_idx, 3, QTableWidgetItem(status))

    def _calibrate_blank_sheet(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Select blank answer sheet", str(Path.home()), SUPPORTED_FILTER)
        if not file_name:
            return
        try:
            pages = load_pages([Path(file_name)], poppler_path=self.settings.poppler_path)
        except Exception as exc:  # pragma: no cover - GUI feedback
            QMessageBox.critical(self, "Calibration Error", f"Failed to load calibration sheet: {exc}")
            return
        if not pages:
            QMessageBox.warning(self, "Calibration", "No pages found in calibration file.")
            return
        page = pages[0]
        self._calibration_page = page
        self.canvas.set_image(page.image)
        self._reset_zoom()
        self._clear_preview_overlays()
        extractions = detect_tables(page.image, page.page_index, str(page.source))
        if not extractions:
            QMessageBox.warning(self, "Calibration", "No tables detected on the calibration sheet.")
            return
        tables_cells: List[List[CellResult]] = []
        for extraction in extractions:
            try:
                region_config = self._resolve_region_assignments(extraction)
                ocr_cells = recognise_table_cells(page.image, extraction, region_config=region_config)
            except TesseractNotFoundError:
                QMessageBox.critical(
                    self,
                    "Tesseract Not Found",
                    "Calibration stopped because Tesseract is unavailable. Install it or set the executable path.",
                )
                return
            extraction.cells = ocr_cells
            tables_cells.append(ocr_cells)
        self._apply_preview_overlays(extractions, tables_cells)
        averages = []
        for cells in tables_cells:
            if not cells:
                continue
            avg_density = sum(cell.ink_density for cell in cells) / len(cells)
            averages.append(avg_density)
        if not averages:
            QMessageBox.information(self, "Calibration", "No cell data available for calibration.")
            return
        baseline = sum(averages) / len(averages)
        self.baseline_density = baseline
        self.settings.baseline_ink_density = baseline
        save_settings(self.settings)
        QMessageBox.information(self, "Calibration", f"Baseline ink density recorded: {baseline:.4f}")
        self.status_bar.showMessage("Calibration baseline saved", 5000)
        reference_extraction = extractions[0]
        if reference_extraction.bounding_box is not None:
            if QMessageBox.question(
                self,
                "Layout Guidance",
                "Would you like to capture the question column and choice rows from this blank sheet?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) == QMessageBox.StandardButton.Yes:
                self._begin_layout_capture(page, reference_extraction)

    def _begin_layout_capture(self, page: LoadedPage, extraction: TableExtraction) -> None:
        if extraction.bounding_box is None:
            QMessageBox.warning(self, "Layout Capture", "Table bounds are unavailable; capture aborted.")
            return
        fallback_labels = self.row_labels_override if self.row_labels_override else ["A", "B", "C", "D"]
        self._reset_layout_capture_state()
        self._reference_table_box = extraction.bounding_box
        self._captured_row_boxes = []
        self._captured_question_box = None
        self._pending_row_capture_labels = list(fallback_labels)
        self._capture_labels_used = list(fallback_labels)
        self.layout_capture_mode = "question"
        self.canvas.set_image(page.image)
        self._reset_zoom()
        self._apply_preview_overlays([extraction], [extraction.cells])
        self.status_bar.showMessage(
            "Draw a rectangle over the question-number column, then release the mouse.",
            8000,
        )
        QMessageBox.information(
            self,
            "Capture Question Column",
            "Use the preview to draw a rectangle tightly around the question-number column, then release.",
        )

    def _handle_layout_selection(self, selection: BoundingBox) -> None:
        if self.layout_capture_mode is None or self._reference_table_box is None:
            return
        relative = self._compute_relative_box(selection, self._reference_table_box)
        if relative is None:
            QMessageBox.warning(
                self,
                "Layout Capture",
                "Selection does not overlap the detected table. Please try again.",
            )
            return
        if self.layout_capture_mode == "question":
            self._captured_question_box = relative
            if not self._pending_row_capture_labels:
                self._complete_layout_capture()
                return
            self.layout_capture_mode = "rows"
            next_label = self._pending_row_capture_labels[0]
            self.status_bar.showMessage(
                f"Draw a rectangle over the '{next_label}' row, then release.",
                8000,
            )
            QMessageBox.information(
                self,
                "Capture Choice Row",
                f"Draw a rectangle tightly around the '{next_label}' row (the answer area for {next_label}).",
            )
            return

        if not self._pending_row_capture_labels:
            return
        label = self._pending_row_capture_labels.pop(0)
        self._captured_row_boxes.append(relative)
        if self._pending_row_capture_labels:
            next_label = self._pending_row_capture_labels[0]
            self.status_bar.showMessage(
                f"Draw a rectangle over the '{next_label}' row, then release.",
                8000,
            )
            QMessageBox.information(
                self,
                "Capture Choice Row",
                f"Draw a rectangle tightly around the '{next_label}' row (the answer area for {next_label}).",
            )
        else:
            self._complete_layout_capture()

    def _complete_layout_capture(self) -> None:
        if self._captured_question_box is None:
            return
        labels_used = list(self._capture_labels_used) if self._capture_labels_used else list(self.row_labels_override)
        if not self._captured_row_boxes or (labels_used and len(self._captured_row_boxes) != len(labels_used)):
            QMessageBox.warning(
                self,
                "Layout Incomplete",
                "Captured rows do not match the expected labels. Please restart the capture process.",
            )
            self._reset_layout_capture_state()
            return

        template_name: Optional[str] = None
        dialog_seed = ""
        while True:
            name, ok = QInputDialog.getText(
                self,
                "Save Layout Template",
                "Enter a name for this layout:",
                text=dialog_seed,
            )
            if not ok:
                self.status_bar.showMessage("Layout capture cancelled.", 5000)
                self._reset_layout_capture_state()
                return
            candidate = name.strip()
            if not candidate:
                QMessageBox.warning(self, "Invalid Name", "Layout name cannot be empty.")
                dialog_seed = name
                continue
            if candidate in self.layout_templates:
                overwrite = QMessageBox.question(
                    self,
                    "Replace Layout",
                    f"A layout named '{candidate}' already exists. Overwrite it?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if overwrite != QMessageBox.StandardButton.Yes:
                    dialog_seed = candidate
                    continue
            template_name = candidate
            break
        if template_name is None:
            self._reset_layout_capture_state()
            return

        layout_data = {
            "question_column": dict(self._captured_question_box),
            "choice_rows": [dict(box) for box in self._captured_row_boxes],
            "row_labels": labels_used,
            "ocr_regions": [dict(box) for box in self.ocr_region_templates],
            "omr_regions": [dict(box) for box in self.omr_region_templates],
        }
        self.layout_templates[template_name] = layout_data
        self.active_layout_name = template_name
        self.question_column_template = dict(self._captured_question_box)
        self.choice_row_templates = [dict(box) for box in self._captured_row_boxes]
        self.row_labels_override = list(labels_used)
        self.settings.row_labels = list(labels_used)
        self.settings.question_column_template = dict(self._captured_question_box)
        self.settings.choice_row_templates = [dict(box) for box in self._captured_row_boxes]
        self.settings.ocr_regions = [dict(box) for box in self.ocr_region_templates]
        self.settings.omr_regions = [dict(box) for box in self.omr_region_templates]
        self.settings.active_layout = template_name
        self.settings.layout_templates = self._serialize_layout_templates()
        save_settings(self.settings)
        self._refresh_layout_selector()
        index = self.layout_selector.findText(template_name)
        if index >= 0:
            self.layout_selector.setCurrentIndex(index)
        self._reset_layout_capture_state()
        self.status_bar.showMessage("Layout template saved. Reload the answer key to apply the new layout.", 8000)
        QMessageBox.information(
            self,
            "Layout Saved",
            f"Layout '{template_name}' has been recorded. Reload the answer key to apply the layout.",
        )

    def _reset_layout_capture_state(self) -> None:
        self.layout_capture_mode = None
        self._pending_row_capture_labels = []
        self._captured_row_boxes = []
        self._captured_question_box = None
        self._reference_table_box = None
        self._capture_labels_used = []

    def _serialize_layout_templates(self) -> List[Dict[str, Any]]:
        return [
            {
                "name": name,
                "question_column": layout.get("question_column", {}),
                "choice_rows": layout.get("choice_rows", []),
                "row_labels": layout.get("row_labels", []),
                "ocr_regions": [dict(box) for box in layout.get("ocr_regions", [])],
                "omr_regions": [dict(box) for box in layout.get("omr_regions", [])],
            }
            for name, layout in sorted(self.layout_templates.items())
        ]

    def _load_layout_templates(self) -> None:
        self.layout_templates = {}
        for entry in getattr(self.settings, "layout_templates", []):
            name = entry.get("name") if isinstance(entry, dict) else None
            question = entry.get("question_column") if isinstance(entry, dict) else None
            rows = entry.get("choice_rows", []) if isinstance(entry, dict) else []
            labels = entry.get("row_labels", []) if isinstance(entry, dict) else []
            if not (name and question and rows):
                continue
            self.layout_templates[name] = {
                "question_column": dict(question),
                "choice_rows": [dict(box) for box in rows],
                "row_labels": list(labels),
                "ocr_regions": [dict(box) for box in entry.get("ocr_regions", [])],
                "omr_regions": [dict(box) for box in entry.get("omr_regions", [])],
            }
        target_layout = None
        if self.settings.active_layout and self.settings.active_layout in self.layout_templates:
            target_layout = self.settings.active_layout
        elif self.layout_templates:
            target_layout = sorted(self.layout_templates.keys())[0]
        if target_layout:
            self._apply_layout(target_layout, persist=False)
        else:
            self.delete_layout_button.setEnabled(False)
        self._refresh_layout_selector()

    def _refresh_layout_selector(self) -> None:
        if not hasattr(self, "layout_selector"):
            return
        self._updating_layout_selector = True
        self.layout_selector.blockSignals(True)
        self.layout_selector.clear()
        for name in sorted(self.layout_templates.keys()):
            self.layout_selector.addItem(name)
        if self.active_layout_name:
            index = self.layout_selector.findText(self.active_layout_name)
            if index >= 0:
                self.layout_selector.setCurrentIndex(index)
        self.layout_selector.blockSignals(False)
        self._updating_layout_selector = False
        self.delete_layout_button.setEnabled(bool(self.layout_templates))

    def _apply_layout(self, name: str, *, persist: bool = True) -> None:
        layout = self.layout_templates.get(name)
        if layout is None:
            return
        self.active_layout_name = name
        self.question_column_template = dict(layout.get("question_column", {}))
        self.choice_row_templates = [dict(box) for box in layout.get("choice_rows", [])]
        self.ocr_region_templates = [dict(box) for box in layout.get("ocr_regions", [])]
        self.omr_region_templates = [dict(box) for box in layout.get("omr_regions", [])]
        labels = list(layout.get("row_labels", []))
        if labels:
            self.row_labels_override = labels
        if persist:
            self.settings.active_layout = name
            self.settings.row_labels = list(self.row_labels_override)
            self.settings.question_column_template = dict(self.question_column_template)
            self.settings.choice_row_templates = [dict(box) for box in self.choice_row_templates]
            self.settings.ocr_regions = [dict(box) for box in self.ocr_region_templates]
            self.settings.omr_regions = [dict(box) for box in self.omr_region_templates]
            save_settings(self.settings)
        self.delete_layout_button.setEnabled(True)
        self._update_region_buttons_state()
        self._refresh_region_overlays()

    def _on_layout_selected(self, index: int) -> None:
        if self._updating_layout_selector or index < 0:
            return
        chosen = self.layout_selector.itemText(index)
        if not chosen or chosen == self.active_layout_name:
            return
        self._apply_layout(chosen)

    def _delete_layout(self) -> None:
        if not self.active_layout_name:
            QMessageBox.information(self, "Delete Layout", "No saved layout is currently active.")
            return
        name = self.active_layout_name
        confirm = QMessageBox.question(
            self,
            "Delete Layout",
            f"Delete layout '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.layout_templates.pop(name, None)
        if self.layout_templates:
            next_name = sorted(self.layout_templates.keys())[0]
            self._apply_layout(next_name, persist=False)
            self.settings.active_layout = next_name
            active_layout = self.layout_templates[next_name]
            self.row_labels_override = list(active_layout.get("row_labels", self.row_labels_override))
            self.question_column_template = dict(active_layout.get("question_column", {}))
            self.choice_row_templates = [dict(box) for box in active_layout.get("choice_rows", [])]
            self.ocr_region_templates = [dict(box) for box in active_layout.get("ocr_regions", [])]
            self.omr_region_templates = [dict(box) for box in active_layout.get("omr_regions", [])]
        else:
            self.active_layout_name = None
            self.question_column_template = None
            self.choice_row_templates = []
            self.ocr_region_templates = []
            self.omr_region_templates = []
            self.settings.active_layout = None
        self.settings.row_labels = list(self.row_labels_override)
        self.settings.question_column_template = (
            dict(self.question_column_template) if self.question_column_template else None
        )
        self.settings.choice_row_templates = [dict(box) for box in self.choice_row_templates]
        self.settings.ocr_regions = [dict(box) for box in self.ocr_region_templates]
        self.settings.omr_regions = [dict(box) for box in self.omr_region_templates]
        self.settings.layout_templates = self._serialize_layout_templates()
        save_settings(self.settings)
        self._refresh_layout_selector()
        self._update_region_buttons_state()
        self._refresh_region_overlays()

    @staticmethod
    def _compute_relative_box(selection: BoundingBox, reference: BoundingBox) -> Optional[Dict[str, float]]:
        ref_right = reference.x + reference.width
        ref_bottom = reference.y + reference.height
        sel_right = selection.x + selection.width
        sel_bottom = selection.y + selection.height
        left = max(selection.x, reference.x)
        top = max(selection.y, reference.y)
        right = min(sel_right, ref_right)
        bottom = min(sel_bottom, ref_bottom)
        if right <= left or bottom <= top or reference.width <= 0 or reference.height <= 0:
            return None
        rel_x = (left - reference.x) / reference.width
        rel_y = (top - reference.y) / reference.height
        rel_w = (right - left) / reference.width
        rel_h = (bottom - top) / reference.height
        return {
            "x": max(0.0, min(1.0, rel_x)),
            "y": max(0.0, min(1.0, rel_y)),
            "width": max(0.0, min(1.0, rel_w)),
            "height": max(0.0, min(1.0, rel_h)),
        }

    @staticmethod
    def _denormalize_relative_box(relative: Dict[str, float], reference: BoundingBox) -> BoundingBox:
        clamp = lambda value: max(0.0, min(1.0, float(value)))
        rel_x = clamp(relative.get("x", 0.0))
        rel_y = clamp(relative.get("y", 0.0))
        rel_w = clamp(relative.get("width", 1.0))
        rel_h = clamp(relative.get("height", 1.0))

        abs_x = reference.x + int(round(rel_x * reference.width))
        abs_y = reference.y + int(round(rel_y * reference.height))
        abs_width = max(1, int(round(rel_w * reference.width)))
        abs_height = max(1, int(round(rel_h * reference.height)))

        max_x = reference.x + reference.width
        max_y = reference.y + reference.height
        if abs_x + abs_width > max_x:
            abs_width = max(1, max_x - abs_x)
        if abs_y + abs_height > max_y:
            abs_height = max(1, max_y - abs_y)

        return BoundingBox(x=abs_x, y=abs_y, width=abs_width, height=abs_height)

    def _configure_row_labels(self) -> None:
        current = ", ".join(self.row_labels_override) if self.row_labels_override else "A, B, C, D"
        text, ok = QInputDialog.getText(
            self,
            "Configure Choice Labels",
            "Enter choice labels in order (comma-separated):",
            text=current,
        )
        if not ok:
            return
        normalized = text.replace(";", ",").replace("\n", ",")
        labels = [part.strip().upper() for part in normalized.split(",") if part.strip()]
        if len(labels) <= 1:
            labels = [part.strip().upper() for part in text.split() if part.strip()]
        if not labels:
            QMessageBox.warning(self, "Invalid Labels", "Please provide at least one label (e.g., A, B, C, D).")
            return
        if self.active_layout_name and self.active_layout_name in self.layout_templates:
            choice_rows = self.layout_templates[self.active_layout_name].get("choice_rows", [])
            if choice_rows and len(labels) != len(choice_rows):
                QMessageBox.warning(
                    self,
                    "Label Mismatch",
                    "The number of labels must match the captured choice rows. Recapture the layout if the sheet format changed.",
                )
                return
            self.layout_templates[self.active_layout_name]["row_labels"] = list(labels)
            self.settings.layout_templates = self._serialize_layout_templates()
        self.row_labels_override = labels
        self.settings.row_labels = labels
        if self.active_layout_name:
            self.settings.active_layout = self.active_layout_name
        self.settings.question_column_template = (
            dict(self.question_column_template) if self.question_column_template else None
        )
        self.settings.choice_row_templates = [dict(box) for box in self.choice_row_templates]
        save_settings(self.settings)
        self.status_bar.showMessage("Choice labels updated", 5000)
        QMessageBox.information(
            self,
            "Layout Updated",
            "Choice labels saved. Reload the answer key to apply them if needed.",
        )

    def _export_results(self) -> None:
        if not self.page_results:
            QMessageBox.information(self, "Export Results", "Process answer sheets before exporting.")
            return
        default_name = str(Path.home() / "mc_marking_results.csv")
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Export Results",
            default_name,
            "CSV Files (*.csv)",
        )
        if not file_name:
            return

        export_path = Path(file_name)
        if export_path.suffix.lower() != ".csv":
            export_path = export_path.with_suffix(".csv")
        if self.answer_key and self.answer_key.rows:
            question_numbers = set(self.answer_key.rows.keys())
        else:
            question_numbers = set()
        for result in self.page_results:
            for answer in result.answers:
                question_numbers.add(answer.question)
        ordered_questions = sorted(question_numbers)

        header = [
            "Source",
            "Page",
            "Correct",
            "Incorrect",
            "Unanswered",
            "Total",
        ]
        for question in ordered_questions:
            header.append(f"Q{question} Answer")
            header.append(f"Q{question} Status")

        try:
            with open(export_path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(header)

                for result in self.page_results:
                    answer_map: Dict[int, PageAnswer] = {answer.question: answer for answer in result.answers}
                    row = [
                        result.source_path.name,
                        result.page_index + 1,
                        result.correct_count,
                        result.incorrect_count,
                        result.unanswered_count,
                        result.total_questions,
                    ]
                    for question in ordered_questions:
                        answer = answer_map.get(question)
                        detected = answer.extracted if answer else ""
                        status: str
                        if answer is None or not answer.extracted:
                            status = "Unanswered"
                        elif answer.is_correct is True:
                            status = "Correct"
                        elif answer.is_correct is False:
                            status = "Incorrect"
                        else:
                            status = "Review"
                        row.extend([detected, status])
                    writer.writerow(row)
        except OSError as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not write file:\n{exc}")
            return

        self.status_bar.showMessage(f"Exported results to {export_path}", 5000)
        QMessageBox.information(self, "Export Complete", f"Student results exported to {export_path}.")

    def _on_selection_changed(self, selection: Optional[BoundingBox]) -> None:
        self.last_selection = selection
        self.use_selection_button.setEnabled(selection is not None)
        if selection is not None:
            self.status_bar.showMessage(
                f"Selection ({selection.width}x{selection.height}) at ({selection.x}, {selection.y})",
                3000,
            )
        if selection is not None and self.region_capture_mode is not None:
            if self._handle_region_selection(selection):
                return
        if selection is not None:
            self._handle_layout_selection(selection)

    def _configure_poppler_path(self) -> None:
        initial_dir = self.settings.poppler_path or str(Path.home())
        directory = QFileDialog.getExistingDirectory(self, "Select Poppler bin directory", initial_dir)
        if not directory:
            return
        self.settings.poppler_path = directory
        save_settings(self.settings)
        self._apply_runtime_paths()
        self.status_bar.showMessage(f"Poppler path set to {directory}", 5000)

    def _configure_tesseract_path(self) -> None:
        initial_dir = self.settings.tesseract_path or str(Path.home())
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select tesseract executable",
            initial_dir,
            "Tesseract Executable (tesseract.exe)",
        )
        if not file_name:
            return
        self.settings.tesseract_path = file_name
        save_settings(self.settings)
        self._apply_runtime_paths()
        self.status_bar.showMessage(f"Tesseract path set to {file_name}", 5000)

    def _apply_runtime_paths(self) -> None:
        if self.settings.tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = self.settings.tesseract_path
        else:
            pytesseract.pytesseract.tesseract_cmd = "tesseract"

    def import_pdf(self) -> None:
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Import PDF", "", SUPPORTED_FILTER)
        if file_path:
            self.status_bar.showMessage(f"Imported: {file_path}")
            is_answer_key = self.answer_key_checkbox.isChecked()
            print(f"File imported: {file_path}, Is first page answer key: {is_answer_key}")


def _build_group_box(title: str, widget: QWidget) -> QGroupBox:
    box = QGroupBox(title)
    layout = QVBoxLayout()
    layout.addWidget(widget)
    box.setLayout(layout)
    return box


def _np_to_qimage(image: np.ndarray) -> QImage:
    if image.ndim != 3 or image.shape[2] != 3:
        raise ValueError("Expected RGB image with shape (h, w, 3)")
    height, width, _ = image.shape
    bytes_per_line = 3 * width
    return QImage(image.data, width, height, bytes_per_line, QImage.Format.Format_RGB888).copy()


def run_app() -> None:
    """Launch the Qt event loop."""
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
