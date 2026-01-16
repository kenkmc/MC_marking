# IMPORTANT: Import easyocr BEFORE PyQt5 to avoid DLL conflicts on Windows
# Try imports for OCR
OCR_ENGINE = None

# Try EasyOCR first (must be before PyQt5 imports)
try:
    import easyocr
    OCR_ENGINE = "easyocr"
    print("EasyOCR loaded successfully")
except (ImportError, OSError, Exception) as e:
    print(f"Warning: EasyOCR not available ({e})")
    # Try Tesseract as fallback
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        OCR_ENGINE = "tesseract"
        print("Using Tesseract")
    except:
        pass

if OCR_ENGINE is None:
    print("Warning: No OCR engine found")

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QFileDialog, QGraphicsView, QGraphicsScene,
    QDialog, QComboBox, QCheckBox, QTextEdit, QGraphicsRectItem,
    QSpinBox, QGroupBox, QTableWidget, QTableWidgetItem, QSplitter,
    QMessageBox, QInputDialog, QScrollArea, QFrame, QSlider,
    QGraphicsPixmapItem, QMenu, QAction
)
from PyQt5.QtGui import QPixmap, QImage, QPen, QBrush, QColor, QPainter, QFont, QWheelEvent, QCursor
from PyQt5.QtCore import Qt, QRectF, QPointF
import fitz  # PyMuPDF for PDF rendering
import sys
import json
import os
import io
import cv2
import numpy as np
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font as XLFont, Alignment, Border, Side

# Mark types
MARK_TYPE_TEXT = "text"      # Text field (e.g., student name, ID)
MARK_TYPE_OPTION = "option"  # Answer option (e.g., A, B, C, D)


def deskew_image(img_array):
    """
    Detect and correct skew in scanned page.
    Returns corrected image and the skew angle.
    """
    # Convert to grayscale if needed
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:
        gray = img_array.copy()
    
    # Apply edge detection
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    
    # Detect lines using Hough transform
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=100, 
                            minLineLength=100, maxLineGap=10)
    
    if lines is None or len(lines) == 0:
        return img_array, 0.0
    
    # Calculate angles of detected lines
    angles = []
    for line in lines:
        x1, y1, x2, y2 = line[0]
        if x2 - x1 != 0:
            angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            # Only consider near-horizontal lines (within 15 degrees)
            if abs(angle) < 15:
                angles.append(angle)
    
    if not angles:
        return img_array, 0.0
    
    # Get median angle (more robust than mean)
    skew_angle = np.median(angles)
    
    # Don't correct very small angles
    if abs(skew_angle) < 0.3:
        return img_array, 0.0
    
    # Rotate image to correct skew
    h, w = img_array.shape[:2]
    center = (w // 2, h // 2)
    rotation_matrix = cv2.getRotationMatrix2D(center, skew_angle, 1.0)
    
    # Calculate new bounding box size
    cos = np.abs(rotation_matrix[0, 0])
    sin = np.abs(rotation_matrix[0, 1])
    new_w = int((h * sin) + (w * cos))
    new_h = int((h * cos) + (w * sin))
    
    # Adjust rotation matrix for new size
    rotation_matrix[0, 2] += (new_w / 2) - center[0]
    rotation_matrix[1, 2] += (new_h / 2) - center[1]
    
    # Apply rotation with white background
    corrected = cv2.warpAffine(img_array, rotation_matrix, (new_w, new_h), 
                               borderMode=cv2.BORDER_CONSTANT, 
                               borderValue=(255, 255, 255) if len(img_array.shape) == 3 else 255)
    
    return corrected, skew_angle


# Modern Style Sheet
STYLE_SHEET = """
QMainWindow {
    background-color: #f0f2f5;
}
QWidget {
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 14px;
}
QGroupBox {
    background-color: white;
    border-radius: 8px;
    margin-top: 1em;
    padding: 15px;
    border: 1px solid #e1e4e8;
    font-weight: bold;
    color: #333;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
    color: #555;
}
QPushButton {
    background-color: #007bff;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    min-height: 20px;
}
QPushButton:hover {
    background-color: #0069d9;
}
QPushButton:pressed {
    background-color: #0062cc;
}
QPushButton:checked {
    background-color: #0056b3;
    border: 2px solid #004085;
}
QPushButton#deleteBtn {
    background-color: #dc3545;
}
QPushButton#deleteBtn:hover {
    background-color: #c82333;
}
QTableWidget {
    border: 1px solid #e1e4e8;
    background-color: white;
    gridline-color: #f0f0f0;
}
QHeaderView::section {
    background-color: #f8f9fa;
    padding: 4px;
    border: 1px solid #e1e4e8;
    font-weight: bold;
}
QLabel#titleLabel {
    font-size: 18px;
    font-weight: bold;
    color: #2c3e50;
    margin-bottom: 10px;
}
"""

class MovablePixmapItem(QGraphicsPixmapItem):
    """A movable pixmap item for the PDF page image."""
    
    def __init__(self, pixmap, parent=None):
        super().__init__(pixmap, parent)
        self.setFlag(QGraphicsPixmapItem.ItemIsMovable, True)
        self.setFlag(QGraphicsPixmapItem.ItemIsSelectable, False)
        self.setAcceptHoverEvents(True)
        self.offset_x = 0
        self.offset_y = 0
        
    def get_offset(self):
        return self.pos().x(), self.pos().y()
        
    def set_offset(self, x, y):
        self.setPos(x, y)


class MarkItem(QGraphicsRectItem):
    """A resizable and movable rectangle for marking areas."""
    
    def __init__(self, x, y, width, height, mark_type=MARK_TYPE_OPTION, 
                 question_num=1, label="", options_count=4, parent=None, view_ref=None):
        super().__init__(x, y, width, height, parent)
        self.mark_type = mark_type
        self.question_num = question_num
        self.label = label
        self.options_count = options_count
        self.view_ref = view_ref
        
        self.setFlag(QGraphicsRectItem.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsRectItem.ItemSendsGeometryChanges, True)
        self.setAcceptHoverEvents(True)
        
        self.update_style()
        
    def update_style(self):
        if self.mark_type == MARK_TYPE_TEXT:
            self.setPen(QPen(QColor(0, 100, 255), 2))
            self.setBrush(QBrush(QColor(0, 100, 255, 50)))
        else:
            self.setPen(QPen(QColor(255, 0, 0), 2))
            self.setBrush(QBrush(QColor(255, 0, 0, 50)))
        
    def paint(self, painter, option, widget):
        super().paint(painter, option, widget)
        rect = self.rect()
        
        if self.mark_type == MARK_TYPE_OPTION:
            # Draw cell divisions for options
            cell_width = rect.width() / self.options_count
            option_labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            
            # Draw vertical dividers
            painter.setPen(QPen(QColor(255, 0, 0, 150), 1, Qt.DashLine))
            for i in range(1, self.options_count):
                x = rect.x() + i * cell_width
                painter.drawLine(int(x), int(rect.y()), int(x), int(rect.y() + rect.height()))
            
            # Draw option labels (A, B, C, D...)
            painter.setPen(QPen(QColor(100, 0, 0)))
            painter.setFont(QFont("Segoe UI", 8))
            for i in range(self.options_count):
                cell_rect = QRectF(rect.x() + i * cell_width, rect.y(), cell_width, rect.height())
                painter.drawText(cell_rect, Qt.AlignCenter, option_labels[i])
            
            # Draw question number at top
            painter.setPen(QPen(Qt.black))
            painter.setFont(QFont("Segoe UI", 9, QFont.Bold))
            display_text = f"Q{self.question_num}"
            if self.label:
                display_text += f" ({self.label})"
            painter.drawText(int(rect.x()), int(rect.y()) - 3, display_text)
        else:
            # Text field - just show the label
            painter.setPen(QPen(Qt.black))
            painter.setFont(QFont("Segoe UI", 9, QFont.Bold))
            display_text = self.label if self.label else f"Field {self.question_num}"
            painter.drawText(rect, Qt.AlignCenter, display_text)
        
    def contextMenuEvent(self, event):
        menu = QMenu()
        rename_action = menu.addAction("Rename" if self.mark_type == MARK_TYPE_TEXT else "Label")
        
        if self.mark_type == MARK_TYPE_OPTION:
            options_action = menu.addAction(f"Set Options (Current: {self.options_count})")
        
        delete_action = menu.addAction("Delete")
        
        action = menu.exec_(event.screenPos())
        
        if action == delete_action:
            if self.view_ref:
                self.view_ref.remove_mark_item(self)
        elif action == rename_action:
            new_label, ok = QInputDialog.getText(None, "Rename", "Enter new label:", text=self.label)
            if ok:
                self.label = new_label
                self.update()
        elif self.mark_type == MARK_TYPE_OPTION and action == options_action:
            count, ok = QInputDialog.getInt(None, "Options Count", "Number of options:", self.options_count, 2, 26)
            if ok:
                self.options_count = count
                self.update()

    def get_data(self):
        scene_rect = self.sceneBoundingRect()
        return {
            "type": self.mark_type,
            "question": self.question_num,
            "label": self.label,
            "options_count": self.options_count,
            "x": scene_rect.x(),
            "y": scene_rect.y(),
            "width": scene_rect.width(),
            "height": scene_rect.height()
        }


class MarkingView(QGraphicsView):
    """Custom graphics view with zoom, marking, and memory."""
    
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        
        # Marking state
        self.marking_mode = False
        self.current_mark_type = MARK_TYPE_OPTION
        self.start_point = None
        self.current_rect = None
        
        # Counters
        self.option_counter = 1
        self.text_counter = 1
        
        # Marks storage
        self.text_marks = []
        self.option_marks = []
        
        # Memory for size
        self.last_option_size = (300, 50) # Default size
        self.last_text_size = (200, 40)   # Default size
        
        # Zoom
        self.zoom_factor = 1.0
        
    def set_marking_mode(self, enabled, mark_type=MARK_TYPE_OPTION):
        self.marking_mode = enabled
        self.current_mark_type = mark_type
        if enabled:
            self.setDragMode(QGraphicsView.NoDrag)
            self.setCursor(Qt.CrossCursor)
        else:
            self.setDragMode(QGraphicsView.ScrollHandDrag)
            self.setCursor(Qt.ArrowCursor)
            
    def remove_mark_item(self, item):
        self.scene().removeItem(item)
        if item in self.text_marks:
            self.text_marks.remove(item)
        if item in self.option_marks:
            self.option_marks.remove(item)
            
    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() == Qt.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                self.zoom_in()
            else:
                self.zoom_out()
            event.accept()
        else:
            super().wheelEvent(event)
            
    def zoom_in(self):
        if self.zoom_factor < 10.0:  # Max zoom limit
            self.zoom_factor *= 1.2
            self.setTransform(QtGui.QTransform().scale(self.zoom_factor, self.zoom_factor))
            
    def zoom_out(self):
        if self.zoom_factor > 0.1:  # Min zoom limit
            self.zoom_factor /= 1.2
            self.setTransform(QtGui.QTransform().scale(self.zoom_factor, self.zoom_factor))
    
    def zoom_reset(self):
        self.zoom_factor = 1.0
        self.setTransform(QtGui.QTransform().scale(1.0, 1.0))
    
    def zoom_fit(self):
        """Fit the entire scene in the view"""
        self.fitInView(self.sceneRect(), Qt.KeepAspectRatio)
        # Update zoom_factor based on current transform
        transform = self.transform()
        self.zoom_factor = transform.m11()  # Get horizontal scale factor
        
    def mousePressEvent(self, event):
        if self.marking_mode and event.button() == Qt.LeftButton:
            self.start_point = self.mapToScene(event.pos())
            
            # Use last known size for this type
            width, height = 0, 0
            
            if self.current_mark_type == MARK_TYPE_TEXT:
                counter = self.text_counter
                # Uncomment to start with default size immediately on click
                # width, height = self.last_text_size 
            else:
                counter = self.option_counter
                # width, height = self.last_option_size
                
            self.current_rect = MarkItem(
                self.start_point.x(), self.start_point.y(), width, height,
                self.current_mark_type, counter, view_ref=self
            )
            self.scene().addItem(self.current_rect)
        else:
            super().mousePressEvent(event)
            
    def mouseMoveEvent(self, event):
        if self.marking_mode and self.current_rect and self.start_point:
            current_pos = self.mapToScene(event.pos())
            rect = QRectF(self.start_point, current_pos).normalized()
            self.current_rect.setRect(rect)
        else:
            super().mouseMoveEvent(event)
            
    def mouseReleaseEvent(self, event):
        if self.marking_mode and self.current_rect:
            rect = self.current_rect.rect()
            
            # If created box is too small, use default/last size
            if rect.width() < 10 or rect.height() < 10:
                if self.current_mark_type == MARK_TYPE_TEXT:
                    w, h = self.last_text_size
                else:
                    w, h = self.last_option_size
                self.current_rect.setRect(self.start_point.x(), self.start_point.y(), w, h)
                rect = self.current_rect.rect() # Update rect
                
            # Save size for next time
            if rect.width() > 10 and rect.height() > 10:
                if self.current_mark_type == MARK_TYPE_TEXT:
                    self.last_text_size = (rect.width(), rect.height())
                    self.text_marks.append(self.current_rect)
                    self.text_counter += 1
                else:
                    self.last_option_size = (rect.width(), rect.height())
                    self.option_marks.append(self.current_rect)
                    self.option_counter += 1
            else:
                self.scene().removeItem(self.current_rect)
            
            self.current_rect = None
            self.start_point = None
        else:
            super().mouseReleaseEvent(event)

    def get_all_marks_data(self):
        marks_data = {
            "text_marks": [],
            "option_marks": []
        }
        for mark in self.text_marks:
            try: marks_data["text_marks"].append(mark.get_data())
            except: continue
        for mark in self.option_marks:
            try: marks_data["option_marks"].append(mark.get_data())
            except: continue
        return marks_data
        
    def load_marks_from_data(self, data):
        # Clear existing first (assumed handled by parent or previous clean)
        pass # Logic handled in OMRSoftware class to avoid duplication


class OMRSoftware(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(STYLE_SHEET)
        
        # OCR Init
        self.ocr_reader = None
        self.ocr_engine_name = OCR_ENGINE
        self.init_ocr()
        
        # Data
        self.pdf_document = None
        self.current_page = 0
        self.page_offsets = {} # Store (x,y) of image per page
        self.marks_data = {} # Full template data
        self.current_pixmap_item = None
        self.answer_key = {}
        self.first_page_key = False
        self.align_reference_gray = None
        self.align_reference_size = None
        
        self.init_ui()
        
    def init_ocr(self):
        if self.ocr_engine_name == "easyocr":
            print("Using EasyOCR")
        elif self.ocr_engine_name == "tesseract":
            print("Using Tesseract")
        else:
            print("No OCR engine found")

    def _prepare_alignment_gray(self, img_np, target_size=None):
        if len(img_np.shape) == 3:
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_np.copy()

        gray = cv2.GaussianBlur(gray, (5, 5), 0)

        if target_size is not None:
            target_w, target_h = target_size
            scale_x = target_w / gray.shape[1]
            scale_y = target_h / gray.shape[0]
            gray = cv2.resize(gray, (target_w, target_h), interpolation=cv2.INTER_AREA)
            return gray, scale_x, scale_y

        # Normalize size to max dimension 800
        max_dim = max(gray.shape[0], gray.shape[1])
        if max_dim > 800:
            scale = 800.0 / max_dim
            gray = cv2.resize(gray, (int(gray.shape[1] * scale), int(gray.shape[0] * scale)), interpolation=cv2.INTER_AREA)
            return gray, scale, scale

        return gray, 1.0, 1.0

    def align_image(self, img_np):
        """
        Align current page to reference by detecting table/frame boundaries.
        Uses edge detection to find the main table frame and align based on its position.
        Returns aligned image, (dx, dy), and confidence score.
        """
        if not hasattr(self, 'check_auto_align') or not self.check_auto_align.isChecked():
            return img_np, (0.0, 0.0), 0.0

        # Find the main table boundaries in current image
        cur_bounds = self._find_table_bounds(img_np)
        
        if cur_bounds is None:
            print("  Auto-align: Cannot detect table boundaries")
            return img_np, (0.0, 0.0), 0.0
        
        # First page becomes the reference
        if self.align_reference_gray is None:
            self.align_reference_bounds = cur_bounds
            self.align_reference_gray = True  # Just mark as initialized
            print(f"  Auto-align: Reference bounds set: {cur_bounds}")
            return img_np, (0.0, 0.0), 1.0

        ref_bounds = self.align_reference_bounds
        
        # Calculate shift needed to align current to reference
        # Align based on top-left corner of detected table
        dx = ref_bounds[0] - cur_bounds[0]  # x shift
        dy = ref_bounds[1] - cur_bounds[1]  # y shift
        
        # Also check if size differs significantly (might indicate wrong detection)
        ref_w = ref_bounds[2] - ref_bounds[0]
        ref_h = ref_bounds[3] - ref_bounds[1]
        cur_w = cur_bounds[2] - cur_bounds[0]
        cur_h = cur_bounds[3] - cur_bounds[1]
        
        size_diff = abs(ref_w - cur_w) / ref_w + abs(ref_h - cur_h) / ref_h
        if size_diff > 0.3:  # More than 30% size difference
            print(f"  Auto-align: Size difference too large ({size_diff:.2%}), skipping")
            return img_np, (0.0, 0.0), 0.5
        
        # Sanity check for extreme shifts
        h, w = img_np.shape[:2]
        if abs(dx) > w * 0.2 or abs(dy) > h * 0.2:
            print(f"  Auto-align: Shift too large (dx={dx:.1f}, dy={dy:.1f}), skipping")
            return img_np, (0.0, 0.0), 0.3
        
        # Skip if shift is negligible
        if abs(dx) < 3 and abs(dy) < 3:
            return img_np, (0.0, 0.0), 1.0
        
        print(f"  Auto-align: Shifting by dx={dx:.1f}, dy={dy:.1f}")
        
        M = np.float32([[1, 0, dx], [0, 1, dy]])
        if len(img_np.shape) == 3:
            border_value = (255, 255, 255)
        else:
            border_value = 255

        aligned = cv2.warpAffine(img_np, M, (w, h), borderMode=cv2.BORDER_CONSTANT, borderValue=border_value)
        return aligned, (dx, dy), 1.0 - size_diff
    
    def _find_table_bounds(self, img_np):
        """
        Find the bounding box of the main table/frame in the image.
        Returns (x1, y1, x2, y2) or None if not found.
        """
        # Convert to grayscale
        if len(img_np.shape) == 3:
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_np.copy()
        
        h, w = gray.shape
        
        # Apply edge detection
        edges = cv2.Canny(gray, 50, 150)
        
        # Apply morphological operations to connect edges
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        edges = cv2.dilate(edges, kernel, iterations=2)
        edges = cv2.erode(edges, kernel, iterations=1)
        
        # Find contours
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            return None
        
        # Find the largest contour that looks like a table (rectangular, large area)
        best_contour = None
        best_area = 0
        min_area = h * w * 0.1  # At least 10% of image
        
        for contour in contours:
            area = cv2.contourArea(contour)
            if area < min_area:
                continue
            
            # Check if roughly rectangular
            peri = cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, 0.02 * peri, True)
            
            # Accept 4-sided polygons or large areas
            if len(approx) >= 4 or area > best_area:
                if area > best_area:
                    best_area = area
                    best_contour = contour
        
        if best_contour is None:
            # Fallback: use projection profile to find table edges
            return self._find_bounds_by_projection(gray)
        
        x, y, bw, bh = cv2.boundingRect(best_contour)
        return (x, y, x + bw, y + bh)
    
    def _find_bounds_by_projection(self, gray):
        """
        Find table bounds using horizontal and vertical projection profiles.
        This works well for scanned documents with clear table borders.
        """
        h, w = gray.shape
        
        # Threshold to binary (invert so lines are white)
        _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
        
        # Horizontal projection (sum along rows)
        h_proj = np.sum(binary, axis=1)
        
        # Vertical projection (sum along columns)  
        v_proj = np.sum(binary, axis=0)
        
        # Find edges using projection threshold
        h_thresh = np.max(h_proj) * 0.1
        v_thresh = np.max(v_proj) * 0.1
        
        # Find top edge
        y1 = 0
        for i in range(h):
            if h_proj[i] > h_thresh:
                y1 = i
                break
        
        # Find bottom edge
        y2 = h - 1
        for i in range(h - 1, -1, -1):
            if h_proj[i] > h_thresh:
                y2 = i
                break
        
        # Find left edge
        x1 = 0
        for i in range(w):
            if v_proj[i] > v_thresh:
                x1 = i
                break
        
        # Find right edge
        x2 = w - 1
        for i in range(w - 1, -1, -1):
            if v_proj[i] > v_thresh:
                x2 = i
                break
        
        # Validate bounds
        if x2 - x1 < w * 0.3 or y2 - y1 < h * 0.3:
            return None
        
        return (x1, y1, x2, y2)

    def detect_filled_option(self, image, options_count=4, save_debug=False):
        """
        Detect which option is filled in a multiple choice bubble area.
        Divides the image into options_count cells and checks which one is filled.
        
        Supports detection of:
        - Dark marks (pencil/pen)
        - Blue marks (blue pen/highlighter)
        - Any colored marks
        
        Args:
            image: PIL Image of the option area
            options_count: Number of options (default 4 for A,B,C,D)
            save_debug: Whether to save debug images
            
        Returns:
            String like "A", "B", "C", "D" or "AB" for multiple selections, or "" if none
        """
        import numpy as np
        import os
        
        img_np = np.array(image)
        
        height, width = img_np.shape[:2]
        cell_width = width // options_count
        
        if cell_width < 5:
            print(f"  Warning: Cell width too small ({cell_width}px)")
            return ""
        
        # Option labels
        option_labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[:options_count]
        
        # Prepare different analysis channels
        if len(img_np.shape) == 3:
            # RGB image - analyze multiple ways
            gray = np.mean(img_np, axis=2)
            
            # Extract color channels
            r_channel = img_np[:, :, 0].astype(float)
            g_channel = img_np[:, :, 1].astype(float)
            b_channel = img_np[:, :, 2].astype(float)
            
            # Calculate color saturation (how "colorful" vs gray)
            max_rgb = np.maximum(np.maximum(r_channel, g_channel), b_channel)
            min_rgb = np.minimum(np.minimum(r_channel, g_channel), b_channel)
            saturation = max_rgb - min_rgb
            
            # Blue detection: high B, low R
            blue_score_img = b_channel - r_channel
            
            has_color = True
        else:
            gray = img_np
            saturation = np.zeros_like(gray)
            blue_score_img = np.zeros_like(gray)
            has_color = False
        
        # Overall statistics
        overall_gray_mean = np.mean(gray)
        overall_sat_mean = np.mean(saturation) if has_color else 0
        
        # Apply contrast enhancement for light mark detection
        # This helps detect very faint pencil marks
        gray_enhanced = gray.copy()
        gray_min, gray_max = np.min(gray), np.max(gray)
        if gray_max > gray_min:
            # Stretch contrast to full range
            gray_enhanced = ((gray - gray_min) / (gray_max - gray_min) * 255).astype(np.float64)
        
        print(f"  Image size: {width}x{height}, {options_count} options, cell width: {cell_width}px")
        print(f"  Overall: gray_mean={overall_gray_mean:.1f}, saturation_mean={overall_sat_mean:.1f}, contrast_range={gray_max-gray_min:.1f}")
        
        # Analyze each cell
        cell_scores = []
        
        for i in range(options_count):
            left = i * cell_width
            right = (i + 1) * cell_width if i < options_count - 1 else width
            
            cell_gray = gray[:, left:right]
            cell_gray_enhanced = gray_enhanced[:, left:right]
            cell_gray_mean = np.mean(cell_gray)
            cell_gray_std = np.std(cell_gray)
            
            # Darkness score (lower mean = darker)
            darkness_score = overall_gray_mean - cell_gray_mean
            
            # Enhanced darkness score using contrast-stretched image
            enhanced_mean = np.mean(cell_gray_enhanced)
            enhanced_overall = np.mean(gray_enhanced)
            enhanced_darkness = enhanced_overall - enhanced_mean
            
            # Local contrast: high std means there's a mark
            local_contrast_score = cell_gray_std / 10.0  # Normalize
            
            # Color-based scores
            if has_color:
                cell_sat = saturation[:, left:right]
                cell_sat_mean = np.mean(cell_sat)
                
                cell_blue = blue_score_img[:, left:right]
                cell_blue_mean = np.mean(cell_blue)
                
                # Saturation difference from overall
                sat_score = cell_sat_mean - overall_sat_mean
            else:
                cell_sat_mean = 0
                cell_blue_mean = 0
                sat_score = 0
            
            # Combined score: weighted sum of different indicators
            # Higher score = more likely to be filled
            # Enhanced scoring for light marks
            combined_score = (
                darkness_score * 1.0 +          # Weight for darkness
                enhanced_darkness * 0.5 +       # Weight for enhanced contrast darkness
                local_contrast_score * 0.3 +    # Weight for local contrast (marks have texture)
                sat_score * 0.5 +               # Weight for saturation (colored marks)
                max(0, cell_blue_mean) * 0.3    # Weight for blue specifically
            )
            
            cell_scores.append({
                'option': option_labels[i],
                'gray_mean': cell_gray_mean,
                'gray_std': cell_gray_std,
                'darkness': darkness_score,
                'enhanced_dark': enhanced_darkness,
                'local_contrast': local_contrast_score,
                'saturation': cell_sat_mean,
                'sat_score': sat_score,
                'blue_score': cell_blue_mean,
                'combined': combined_score
            })
            
            print(f"    Option {option_labels[i]}: gray={cell_gray_mean:.1f}, dark={darkness_score:.1f}, enh_dark={enhanced_darkness:.1f}, contrast={local_contrast_score:.1f}, combined={combined_score:.1f}")
        
        # Save debug image with cell divisions and scores
        if save_debug:
            from PIL import ImageDraw, ImageFont
            debug_dir = "debug_crops"
            os.makedirs(debug_dir, exist_ok=True)
            import time
            
            debug_img = image.copy()
            draw = ImageDraw.Draw(debug_img)
            
            # Draw vertical lines to show cell divisions
            for i in range(1, options_count):
                x = i * cell_width
                draw.line([(x, 0), (x, height)], fill=(255, 0, 0), width=2)
            
            # Draw scores on each cell
            for i, score in enumerate(cell_scores):
                x = i * cell_width + 2
                draw.text((x, 2), f"{score['combined']:.0f}", fill=(255, 0, 0))
            
            debug_path = os.path.join(debug_dir, f"option_{int(time.time()*1000)}.png")
            debug_img.save(debug_path)
            print(f"  Saved debug image: {debug_path}")
        
        # Determine which option(s) are filled using combined score
        filled_options = []
        
        if cell_scores:
            max_combined = max(s['combined'] for s in cell_scores)
            min_combined = min(s['combined'] for s in cell_scores)
            score_range = max_combined - min_combined
            
            # Dynamic threshold based on score range
            # If there's clear variation, use relative threshold
            # Otherwise, use absolute threshold
            
            print(f"  Score range: {min_combined:.1f} to {max_combined:.1f} (range={score_range:.1f})")
            
            # Adaptive thresholding for light marks
            # Much lower thresholds to detect very faint pencil marks
            
            # Method 1: If there's clear variation, use the highest scoring option
            if score_range > 1.5:  # Even small variation matters (lowered from 2)
                # Find the option with highest combined score
                max_score_option = max(cell_scores, key=lambda s: s['combined'])
                
                # Pick options that are close to the maximum (within 30% of range from top)
                threshold_combined = max_combined - score_range * 0.35
                
                # Very low minimum absolute threshold for light marks
                min_absolute = 1.0  # Very low to catch faint marks
                
                for score in cell_scores:
                    if score['combined'] >= threshold_combined and score['combined'] >= min_absolute:
                        filled_options.append(score['option'])
                
                # If still no detection but there's a clear winner, pick it
                if not filled_options and max_combined > 0.5:
                    filled_options.append(max_score_option['option'])
                        
                print(f"  Threshold (relative): {threshold_combined:.1f}, min_absolute: {min_absolute}")
            else:
                # No clear variation - use standard deviation analysis
                # Check if any option has significantly lower gray mean (darker)
                gray_means = [s['gray_mean'] for s in cell_scores]
                gray_std_overall = np.std(gray_means) if len(gray_means) > 1 else 0
                
                if gray_std_overall > 1:  # There's some variation in gray levels
                    min_gray = min(gray_means)
                    for score in cell_scores:
                        # Pick options that are darker than average
                        if score['gray_mean'] <= min_gray + gray_std_overall * 0.5:
                            if score['darkness'] > 0:  # Must be at least slightly darker
                                filled_options.append(score['option'])
                    print(f"  Using gray std analysis: std={gray_std_overall:.1f}")
                else:
                    # Fallback: very low absolute threshold
                    threshold_combined = 1.5  # Very low for light marks
                    
                    for score in cell_scores:
                        if score['combined'] > threshold_combined:
                            filled_options.append(score['option'])
                            
                    print(f"  Threshold (absolute): {threshold_combined}")
                        
                print(f"  Threshold (absolute): {threshold_combined}")
        
        result = "".join(filled_options)
        print(f"  Detected filled option(s): {result if result else '(none)'}")
        return result

    def get_ocr_result(self, image, save_debug=False):
        """Perform OCR on the given PIL image and return text with confidence info."""
        import numpy as np
        
        # Debug: Save cropped image to see what's being recognized
        if save_debug:
            import os
            debug_dir = "debug_crops"
            os.makedirs(debug_dir, exist_ok=True)
            import time
            debug_path = os.path.join(debug_dir, f"crop_{int(time.time()*1000)}.png")
            image.save(debug_path)
            print(f"  Saved debug image: {debug_path}")
        
        # Check if image is valid
        img_np = np.array(image)
        print(f"  Image shape: {img_np.shape}, dtype: {img_np.dtype}")
        
        if img_np.size == 0:
            print("  ERROR: Empty image!")
            return "[Empty Image]"
        
        if self.ocr_engine_name == "easyocr":
            if self.ocr_reader is None:
                import easyocr
                # Initialize for English and Traditional Chinese
                print("  Initializing EasyOCR reader (this may take a moment)...")
                self.ocr_reader = easyocr.Reader(['en', 'ch_tra'], verbose=False) 
            
            # Get detailed results with confidence
            result = self.ocr_reader.readtext(img_np, detail=1)
            
            if not result:
                print("  EasyOCR: No text detected")
                return ""
            
            texts = []
            for detection in result:
                bbox, text, confidence = detection
                print(f"  EasyOCR detected: '{text}' (confidence: {confidence:.2%})")
                texts.append(text)
            
            return " ".join(texts)
        
        elif self.ocr_engine_name == "tesseract":
            import pytesseract
            # Default to eng+chi_tra
            try:
                text = pytesseract.image_to_string(image, lang='eng+chi_tra').strip()
                print(f"  Tesseract detected: '{text}'")
                return text
            except:
                text = pytesseract.image_to_string(image, lang='eng').strip()
                print(f"  Tesseract detected: '{text}'")
                return text
        
        return "OCR Error: No Engine"

    def init_ui(self):
        self.setWindowTitle("OMR Marking Software - Modern Edition")
        self.setGeometry(100, 100, 1400, 850)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        
        # --- Left Panel (Controls) ---
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setFixedWidth(380)
        
        left_content = QWidget()
        left_layout = QVBoxLayout(left_content)
        left_layout.setSpacing(15)
        
        # Title
        title = QLabel("OMR Master")
        title.setObjectName("titleLabel")
        left_layout.addWidget(title)
        
        # File Import
        file_grp = QGroupBox("1. File & Setup")
        f_layout = QVBoxLayout(file_grp)
        
        btn_import = QPushButton("Import PDF")
        btn_import.clicked.connect(self.import_pdf)
        f_layout.addWidget(btn_import)
        
        self.check_first_key = QCheckBox("First page is Answer Key")
        self.check_first_key.stateChanged.connect(lambda s: setattr(self, 'first_page_key', s == Qt.Checked))
        f_layout.addWidget(self.check_first_key)
        
        self.check_auto_deskew = QCheckBox("Auto-correct page skew")
        self.check_auto_deskew.setChecked(True)  # Enable by default
        f_layout.addWidget(self.check_auto_deskew)

        self.check_auto_align = QCheckBox("Auto-align pages (shift)")
        self.check_auto_align.setChecked(True)  # Enable by default
        f_layout.addWidget(self.check_auto_align)
        
        left_layout.addWidget(file_grp)
        
        # Marking Tools
        mark_grp = QGroupBox("2. Marking Tools")
        m_layout = QVBoxLayout(mark_grp)
        
        row1 = QHBoxLayout()
        self.btn_mark_text = QPushButton("Mark Text Field")
        self.btn_mark_text.setCheckable(True)
        self.btn_mark_text.clicked.connect(lambda: self.set_marking(MARK_TYPE_TEXT))
        row1.addWidget(self.btn_mark_text)
        
        self.btn_mark_option = QPushButton("Mark Options")
        self.btn_mark_option.setCheckable(True)
        self.btn_mark_option.clicked.connect(lambda: self.set_marking(MARK_TYPE_OPTION))
        row1.addWidget(self.btn_mark_option)
        m_layout.addLayout(row1)
        
        m_layout.addWidget(QLabel("Right-click marks to Rename/Delete/Config"))
        
        row2 = QHBoxLayout()
        btn_clear = QPushButton("Clear All")
        btn_clear.setObjectName("deleteBtn")
        btn_clear.clicked.connect(self.clear_all_marks)
        row2.addWidget(btn_clear)
        m_layout.addLayout(row2)
        
        row3 = QHBoxLayout()
        btn_import_templ = QPushButton("Load Template")
        btn_import_templ.clicked.connect(self.import_template)
        btn_export_templ = QPushButton("Save Template")
        btn_export_templ.clicked.connect(self.export_template)
        row3.addWidget(btn_import_templ)
        row3.addWidget(btn_export_templ)
        m_layout.addLayout(row3)
        
        left_layout.addWidget(mark_grp)
        
        # Processing
        proc_grp = QGroupBox("3. Processing")
        p_layout = QVBoxLayout(proc_grp)
        
        lbl_ocr = QLabel(f"OCR Status: {self.ocr_engine_name if self.ocr_engine_name else 'Not Available'}")
        p_layout.addWidget(lbl_ocr)
        
        btn_process = QPushButton("Recognize All Pages")
        btn_process.clicked.connect(self.run_recognition_all)
        p_layout.addWidget(btn_process)
        
        btn_export = QPushButton("Export to Excel")
        btn_export.clicked.connect(self.export_excel)
        p_layout.addWidget(btn_export)
        
        btn_export_img = QPushButton("Export Images with Answers")
        btn_export_img.clicked.connect(self.export_images)
        p_layout.addWidget(btn_export_img)
        
        left_layout.addWidget(proc_grp)
        
        left_layout.addStretch()
        left_scroll.setWidget(left_content)
        layout.addWidget(left_scroll)
        
        # --- Center (Preview) ---
        center_layout = QVBoxLayout()
        
        # Toolbar with navigation and zoom
        nav_layout = QHBoxLayout()
        btn_prev = QPushButton("◀ Prev")
        btn_prev.clicked.connect(self.prev_page)
        btn_next = QPushButton("Next ▶")
        btn_next.clicked.connect(self.next_page)
        self.lbl_page = QLabel("Page: 0/0")
        
        # Zoom controls
        btn_zoom_in = QPushButton("🔍+")
        btn_zoom_in.setFixedWidth(40)
        btn_zoom_in.setToolTip("Zoom In (Ctrl+Scroll Up)")
        btn_zoom_in.clicked.connect(lambda: self.view.zoom_in())
        
        btn_zoom_out = QPushButton("🔍-")
        btn_zoom_out.setFixedWidth(40)
        btn_zoom_out.setToolTip("Zoom Out (Ctrl+Scroll Down)")
        btn_zoom_out.clicked.connect(lambda: self.view.zoom_out())
        
        btn_zoom_reset = QPushButton("100%")
        btn_zoom_reset.setFixedWidth(50)
        btn_zoom_reset.setToolTip("Reset Zoom")
        btn_zoom_reset.clicked.connect(lambda: self.view.zoom_reset())
        
        btn_zoom_fit = QPushButton("Fit")
        btn_zoom_fit.setFixedWidth(40)
        btn_zoom_fit.setToolTip("Fit to Window")
        btn_zoom_fit.clicked.connect(lambda: self.view.zoom_fit())
        
        nav_layout.addWidget(btn_prev)
        nav_layout.addWidget(self.lbl_page)
        nav_layout.addWidget(btn_next)
        nav_layout.addStretch()
        nav_layout.addWidget(QLabel("Zoom:"))
        nav_layout.addWidget(btn_zoom_out)
        nav_layout.addWidget(btn_zoom_reset)
        nav_layout.addWidget(btn_zoom_in)
        nav_layout.addWidget(btn_zoom_fit)
        
        center_layout.addLayout(nav_layout)
        
        self.scene = QGraphicsScene()
        self.view = MarkingView(self.scene)
        center_layout.addWidget(self.view)
        
        layout.addLayout(center_layout, stretch=2)
        
        # --- Right (Results) ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_widget.setFixedWidth(350)
        
        right_layout.addWidget(QLabel("<b>Results & Answer Key</b>"))
        
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Q", "Detected", "Correct", "Points"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.cellChanged.connect(self.on_table_edit)
        right_layout.addWidget(self.table)
        
        self.lbl_score = QLabel("Total: 0")
        right_layout.addWidget(self.lbl_score)
        
        layout.addWidget(right_widget)

    def set_marking(self, mtype):
        if mtype == MARK_TYPE_TEXT:
            is_checked = self.btn_mark_text.isChecked()
            self.btn_mark_option.setChecked(False)
            self.view.set_marking_mode(is_checked, MARK_TYPE_TEXT)
        else:
            is_checked = self.btn_mark_option.isChecked()
            self.btn_mark_text.setChecked(False)
            self.view.set_marking_mode(is_checked, MARK_TYPE_OPTION)

    def import_pdf(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if fname:
            try:
                self.pdf_path = fname
                self.pdf_document = fitz.open(fname)
                self.current_page = 0
                self.align_reference_gray = None
                self.align_reference_size = None
                self.load_page(0)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _get_pdf_prefix(self):
        if hasattr(self, 'pdf_path') and self.pdf_path:
            base = os.path.splitext(os.path.basename(self.pdf_path))[0]
            return base
        return "export"

    def _get_timestamp(self):
        return QtCore.QDateTime.currentDateTime().toString("yyyyMMdd_HHmmss")

    def load_page(self, p_idx):
        if not self.pdf_document: return
        
        # Save current image offset
        if self.current_pixmap_item:
            self.page_offsets[self.current_page] = self.current_pixmap_item.get_offset()
            
        self.current_page = p_idx
        self.lbl_page.setText(f"Page: {p_idx+1}/{len(self.pdf_document)}")
        
        # Render PDF
        page = self.pdf_document[p_idx]
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        
        # Remove only the pixmap item, not the marks
        if self.current_pixmap_item is not None:
            self.scene.removeItem(self.current_pixmap_item)
        
        # Add Image
        pix_item = QPixmap.fromImage(img)
        self.current_pixmap_item = MovablePixmapItem(pix_item)
        
        # Restore offset
        if p_idx in self.page_offsets:
            off = self.page_offsets[p_idx]
            self.current_pixmap_item.setPos(off[0], off[1])
            
        self.scene.addItem(self.current_pixmap_item)
        # Move pixmap to back so marks are visible on top
        self.current_pixmap_item.setZValue(-1)
        self.view.setSceneRect(QRectF(0, 0, pix.width, pix.height))
        
        # Reset zoom when changing pages to ensure marks align correctly
        self.view.zoom_reset()
        
        # Ensure marks are in the scene
        for m in self.view.text_marks + self.view.option_marks:
            if m.scene() is None:
                self.scene.addItem(m)

        # Update table for this page result if available
        self.update_result_table()

    def prev_page(self):
        if self.current_page > 0:
            self.load_page(self.current_page - 1)

    def next_page(self):
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.load_page(self.current_page + 1)

    def clear_all_marks(self):
        for m in self.view.text_marks + self.view.option_marks:
            self.scene.removeItem(m)
        self.view.text_marks.clear()
        self.view.option_marks.clear()
        self.view.text_counter = 1
        self.view.option_counter = 1

    def export_template(self):
        data = self.view.get_all_marks_data()
        fname, _ = QFileDialog.getSaveFileName(self, "Save Template", "", "JSON (*.json)")
        if fname:
            with open(fname, 'w') as f:
                json.dump(data, f, indent=2)

    def import_template(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Load Template", "", "JSON (*.json)")
        if fname:
            with open(fname, 'r') as f:
                data = json.load(f)
            self.clear_all_marks()
            
            for m in data.get("text_marks", []):
                item = MarkItem(m['x'], m['y'], m['width'], m['height'], MARK_TYPE_TEXT, m['question'], m['label'], view_ref=self.view)
                self.view.text_marks.append(item)
                self.scene.addItem(item)
                self.view.text_counter = max(self.view.text_counter, m['question'] + 1)
                
            for m in data.get("option_marks", []):
                item = MarkItem(m['x'], m['y'], m['width'], m['height'], MARK_TYPE_OPTION, m['question'], m['label'], m.get('options_count', 4), view_ref=self.view)
                self.view.option_marks.append(item)
                self.scene.addItem(item)
                self.view.option_counter = max(self.view.option_counter, m['question'] + 1)

    def run_recognition_all(self):
        if not self.pdf_document: 
            QMessageBox.warning(self, "Warning", "No PDF loaded")
            return
        if not self.view.option_marks and not self.view.text_marks:
            QMessageBox.warning(self, "Warning", "No marks defined. Please mark regions first.")
            return
            
        self.results = {}
        
        # Save current page's image offset before processing
        if self.current_pixmap_item:
            self.page_offsets[self.current_page] = self.current_pixmap_item.get_offset()
        
        # Progress Dialog
        progress = QtWidgets.QProgressDialog("Recognizing...", "Cancel", 0, len(self.pdf_document), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()
        
        # Get marks data once (they are the same for all pages)
        # Marks are in scene coordinates. We need to map them relative to image position.
        
        for p_idx in range(len(self.pdf_document)):
            QtWidgets.QApplication.processEvents()
            if progress.wasCanceled(): 
                break
            progress.setValue(p_idx)
            progress.setLabelText(f"Recognizing page {p_idx + 1} of {len(self.pdf_document)}...")
            
            # Render page
            page = self.pdf_document[p_idx]
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Apply auto-deskew if enabled
            skew_angle = 0.0
            if self.check_auto_deskew.isChecked():
                img_np = np.array(img_pil)
                img_corrected, skew_angle = deskew_image(img_np)
                if skew_angle != 0.0:
                    print(f"Page {p_idx + 1}: Corrected skew angle: {skew_angle:.2f}°")
                    img_pil = Image.fromarray(img_corrected)

            # Apply auto-align (shift) if enabled
            if self.check_auto_align.isChecked():
                img_np = np.array(img_pil)
                img_aligned, (dx, dy), response = self.align_image(img_np)
                if dx != 0.0 or dy != 0.0:
                    print(f"Page {p_idx + 1}: Aligned shift dx={dx:.1f}, dy={dy:.1f} (score={response:.3f})")
                    img_pil = Image.fromarray(img_aligned)
            
            # Get Image Offset for this page (where the image was positioned in the scene)
            # If user moved the image, marks are relative to scene origin (0,0)
            # Image is at (off_x, off_y), so to get image-relative coords:
            # image_x = scene_x - off_x
            off_x, off_y = self.page_offsets.get(p_idx, (0, 0))
            
            page_res = {
                "options": {},
                "text": {}
            }
            
            # Helper to process a list of marks
            def process_marks(marks_list, target_dict):
                for mark in marks_list:
                    rect = mark.sceneBoundingRect()
                    # Convert scene coordinates to image coordinates
                    # The image is positioned at (off_x, off_y) in the scene
                    # So image coordinate = scene coordinate - image offset
                    img_x = rect.x() - off_x
                    img_y = rect.y() - off_y
                    
                    print(f"Mark Q{mark.question_num}: scene=({rect.x():.0f},{rect.y():.0f}), offset=({off_x:.0f},{off_y:.0f}), img=({img_x:.0f},{img_y:.0f}), size=({rect.width():.0f}x{rect.height():.0f})")
                    
                    # Ensure crop is within image bounds
                    left = max(0, int(img_x))
                    top = max(0, int(img_y))
                    right = min(img_pil.width, int(img_x + rect.width()))
                    bottom = min(img_pil.height, int(img_y + rect.height()))
                    
                    print(f"  Crop: ({left},{top})-({right},{bottom}), img size: {img_pil.width}x{img_pil.height}")
                    
                    if right > left and bottom > top:
                        crop = img_pil.crop((left, top, right, bottom))
                        
                        if mark.mark_type == MARK_TYPE_OPTION:
                            # Use bubble detection for options
                            text = self.detect_filled_option(crop, mark.options_count, save_debug=True)
                        else:
                            # Use OCR for text fields
                            text = self.get_ocr_result(crop, save_debug=True)
                    else:
                        text = f"[Out of bounds]"
                        print(f"  Out of bounds!")
                    
                    if mark.mark_type == MARK_TYPE_OPTION:
                        target_dict[mark.question_num] = text
                    else:
                        # For text fields, use label as key if exists, else "Field X"
                        key = mark.label if mark.label else f"Field {mark.question_num}"
                        target_dict[key] = text
            
            process_marks(self.view.option_marks, page_res["options"])
            process_marks(self.view.text_marks, page_res["text"])
            
            # Store
            self.results[p_idx] = page_res
            
        progress.setValue(len(self.pdf_document))
        progress.close()
        
        # Show summary
        total_pages = len(self.results)
        total_options = sum(len(r.get("options", {})) for r in self.results.values())
        QMessageBox.information(self, "Recognition Complete", 
            f"Processed {total_pages} pages.\nRecognized {total_options} option fields.")
        
        self.update_result_table()

        # Build Answer Key if needed
        if self.first_page_key and 0 in self.results:
            self.answer_key = self.results[0]["options"]
            # Refresh to show scores
            self.update_result_table()


    def update_result_table(self):
        # Display results for CURRENT page
        if self.current_page not in getattr(self, 'results', {}):
            return
            
        page_res = self.results[self.current_page]
        # structure: {"options": {1: "A", ...}, "text": {"Name": "John", ...}}
        
        opts = page_res.get("options", {})
        texts = page_res.get("text", {})
        
        self.table.setRowCount(len(texts) + len(opts))
        
        current_row = 0
        
        # 1. Text Fields
        for key, val in texts.items():
            self.table.setItem(current_row, 0, QTableWidgetItem(str(key)))
            self.table.setItem(current_row, 1, QTableWidgetItem(str(val)))
            self.table.setItem(current_row, 2, QTableWidgetItem("-")) # No correct answer for info
            self.table.setItem(current_row, 3, QTableWidgetItem("-")) # No points
            
            # Grey out key/points
            self.table.item(current_row, 2).setFlags(Qt.ItemIsEnabled)
            self.table.item(current_row, 3).setFlags(Qt.ItemIsEnabled)
            current_row += 1
            
        # 2. Options
        sorted_qs = sorted(opts.keys())
        total_score = 0
        
        for q_num in sorted_qs:
            detected = opts[q_num]
            correct = self.answer_key.get(q_num, "")
            
            # Normalize for comparison
            is_correct = False
            if correct and detected:
                # remove spaces, lowercase
                d_clean = "".join(detected.split()).lower()
                c_clean = "".join(str(correct).split()).lower()
                if d_clean == c_clean: is_correct = True
            
            points = 1 if is_correct else 0
            if is_correct: total_score += 1
            
            self.table.setItem(current_row, 0, QTableWidgetItem(f"Q{q_num}"))
            self.table.setItem(current_row, 1, QTableWidgetItem(str(detected)))
            self.table.setItem(current_row, 2, QTableWidgetItem(str(correct)))
            self.table.setItem(current_row, 3, QTableWidgetItem(str(points)))
            
            # Color code
            if correct:
                color = QColor("#d4edda") if is_correct else QColor("#f8d7da")
                self.table.item(current_row, 1).setBackground(color)
            
            current_row += 1
        
        self.lbl_score.setText(f"Page Score: {total_score}")

    def on_table_edit(self, row, col):
        if col == 2: # Correct Answer column
            item_header = self.table.item(row, 0).text()
            if item_header.startswith("Q"):
                try:
                    q_txt = item_header.replace("Q", "")
                    q_num = int(q_txt)
                    new_ans = self.table.item(row, 2).text()
                    self.answer_key[q_num] = new_ans
                    # For immediate feedback we could call update_result_table(),
                    # but be careful of infinite recursion if we were generating signals.
                    # Since we are reacting to user edit, it's fine.
                except: pass
            else:
                # Text field, reset if user tries to edit key
                self.table.item(row, 2).setText("-")

    def export_excel(self):
        if not hasattr(self, 'results'): return
        prefix = self._get_pdf_prefix()
        timestamp = self._get_timestamp()
        default_name = f"{prefix}_{timestamp}.xlsx"
        fname, _ = QFileDialog.getSaveFileName(self, "Export Excel", default_name, "Excel (*.xlsx)")
        if not fname: return

        # Ensure filename includes prefix and timestamp
        base_dir = os.path.dirname(fname)
        fname = os.path.join(base_dir, default_name)
        
        # Import styling for highlighting empty cells
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        # Define styles
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow for empty answers
        orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # Orange for multiple selections
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green for stats header
        header_font = Font(bold=True)
        center_align = Alignment(horizontal='center')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        wb = Workbook()
        ws = wb.active
        ws.title = "OMR Results"
        
        # Gather all headers
        all_qs = set()
        all_texts = set()
        
        for p_res in self.results.values():
            all_qs.update(p_res.get("options", {}).keys())
            all_texts.update(p_res.get("text", {}).keys())
            
        sorted_qs = sorted(list(all_qs))
        sorted_texts = sorted(list(all_texts))
        
        # Calculate column positions
        # Column A = Page, then text fields, then questions, then Score
        text_start_col = 2  # B
        q_start_col = text_start_col + len(sorted_texts)  # After text fields
        score_col = q_start_col + len(sorted_qs)  # After questions
        
        # Headers: Page, [Text Fields], [Questions], Total Score
        headers = ["Page"] + sorted_texts + [f"Q{q}" for q in sorted_qs] + ["Score"]
        ws.append(headers)
        
        # Key Row (row 2)
        key_row = ["Key"] + [""] * len(sorted_texts)
        for q in sorted_qs: key_row.append(self.answer_key.get(q, ""))
        key_row.append("")  # No score for key row
        ws.append(key_row)
        
        # Data rows start from row 3
        data_row_num = 3
        first_data_row = 3
        empty_cells = []  # Track cells with empty answers for highlighting
        multiple_cells = []  # Track cells with multiple selections for highlighting
        
        for p_idx, res in self.results.items():
            if self.first_page_key and p_idx == 0: continue
            
            row = [p_idx + 1]
            
            # Text Fields
            texts = res.get("text", {})
            for t_key in sorted_texts:
                row.append(texts.get(t_key, ""))
                
            # Questions - track empty answers and multiple selections
            opts = res.get("options", {})
            for q_idx, q in enumerate(sorted_qs):
                val = opts.get(q, "")
                row.append(val)
                col_letter = get_column_letter(q_start_col + q_idx)
                # Track empty answers for highlighting
                if val == "" or val is None:
                    empty_cells.append(f"{col_letter}{data_row_num}")
                # Track multiple selections (e.g., "AB", "BC") for highlighting
                elif len(str(val)) > 1:
                    multiple_cells.append(f"{col_letter}{data_row_num}")
            
            # Score formula using SUMPRODUCT to compare each answer with key row
            # Formula: =SUMPRODUCT((C3:Z3=C$2:Z$2)*1) where C:Z are question columns
            if sorted_qs:
                first_q_col = get_column_letter(q_start_col)
                last_q_col = get_column_letter(q_start_col + len(sorted_qs) - 1)
                score_formula = f'=SUMPRODUCT(({first_q_col}{data_row_num}:{last_q_col}{data_row_num}={first_q_col}$2:{last_q_col}$2)*1)'
                row.append(score_formula)
            else:
                row.append(0)
            
            ws.append(row)
            data_row_num += 1
        
        last_data_row = data_row_num - 1
        
        # Apply yellow highlighting to empty answer cells
        for cell_ref in empty_cells:
            ws[cell_ref].fill = yellow_fill
        
        # Apply orange highlighting to multiple selection cells
        for cell_ref in multiple_cells:
            ws[cell_ref].fill = orange_fill
        
        # Add Statistics Row - Percentage correct per question
        if sorted_qs and last_data_row >= first_data_row:
            stats_row_num = data_row_num + 1  # Skip one row
            
            # Stats header row
            ws.cell(row=stats_row_num, column=1, value="% Correct").fill = green_fill
            ws.cell(row=stats_row_num, column=1).font = header_font
            
            # Add percentage formula for each question
            # Formula: =COUNTIF(col_range, key_cell) / COUNT of non-empty cells * 100
            for q_idx, q in enumerate(sorted_qs):
                col_num = q_start_col + q_idx
                col_letter = get_column_letter(col_num)
                
                # Calculate % correct: count matches with key / total responses
                # =COUNTIF(C3:C10, C2) / COUNTA(C3:C10) * 100
                # COUNTA counts non-empty cells
                data_range = f"{col_letter}{first_data_row}:{col_letter}{last_data_row}"
                key_cell = f"{col_letter}$2"
                
                percent_formula = f'=IF(COUNTA({data_range})>0, COUNTIF({data_range},{key_cell})/COUNTA({data_range})*100, 0)'
                
                cell = ws.cell(row=stats_row_num, column=col_num, value=percent_formula)
                cell.fill = green_fill
                cell.alignment = center_align
                cell.number_format = '0.0"%"'
            
            # Average % correct in Score column
            if sorted_qs:
                first_q_col = get_column_letter(q_start_col)
                last_q_col = get_column_letter(q_start_col + len(sorted_qs) - 1)
                avg_formula = f'=AVERAGE({first_q_col}{stats_row_num}:{last_q_col}{stats_row_num})'
                cell = ws.cell(row=stats_row_num, column=score_col, value=avg_formula)
                cell.fill = green_fill
                cell.alignment = center_align
                cell.number_format = '0.0"%"'
        
        # Style header row
        for col in range(1, len(headers) + 1):
            ws.cell(row=1, column=col).font = header_font
            ws.cell(row=1, column=col).alignment = center_align
            
        wb.save(fname)
        
        empty_count = len(empty_cells)
        multiple_count = len(multiple_cells)
        msg = "Export complete! Score is calculated by formula.\n"
        if empty_count > 0:
            msg += f"\n⚠️ {empty_count} empty answers highlighted in yellow."
        if multiple_count > 0:
            msg += f"\n🔶 {multiple_count} multiple selections highlighted in orange."
        msg += "\n\n📊 Per-question % correct statistics added at bottom."
        QMessageBox.information(self, "Done", msg)
    
    def export_images(self):
        """Export scanned pages as images with answer overlay (red dots for correct answers)"""
        if not hasattr(self, 'pdf_document') or self.pdf_document is None:
            QMessageBox.warning(self, "Error", "No PDF loaded")
            return

        if not hasattr(self, 'view') or not self.view.option_marks:
            QMessageBox.warning(self, "Error", "No option marks found")
            return
        
        # Use PDF directory as default location
        prefix = self._get_pdf_prefix()
        timestamp = self._get_timestamp()
        if hasattr(self, 'pdf_path') and self.pdf_path:
            parent_folder = os.path.dirname(self.pdf_path)
        else:
            parent_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if not parent_folder: return
        
        folder = os.path.join(parent_folder, f"{prefix}_{timestamp}")
        os.makedirs(folder, exist_ok=True)
        
        from PyQt5.QtWidgets import QProgressDialog
        progress = QProgressDialog("Exporting images...", "Cancel", 0, len(self.pdf_document), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        
        for page_idx in range(len(self.pdf_document)):
            if progress.wasCanceled(): break
            progress.setValue(page_idx)
            QtWidgets.QApplication.processEvents()
            
            # Render page at 2x scale
            page = self.pdf_document[page_idx]
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)

            # Convert to numpy array for processing
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_np = np.array(img_pil)

            # Apply auto-deskew if enabled
            if self.check_auto_deskew.isChecked():
                img_np, skew_angle = deskew_image(img_np)
                if skew_angle != 0.0:
                    print(f"Export page {page_idx + 1}: Corrected skew angle: {skew_angle:.2f}°")

            # Apply auto-align (shift) if enabled
            if self.check_auto_align.isChecked():
                img_np, (dx, dy), response = self.align_image(img_np)
                if dx != 0.0 or dy != 0.0:
                    print(f"Export page {page_idx + 1}: Aligned shift dx={dx:.1f}, dy={dy:.1f} (score={response:.3f})")

            # Convert to QImage for drawing
            h, w = img_np.shape[:2]
            qimg = QImage(img_np.data, w, h, img_np.strides[0], QImage.Format_RGB888).copy()
            
            # Create painter to draw overlay
            painter = QPainter(qimg)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # Get results for this page
            page_results = self.results.get(page_idx, {}) if hasattr(self, 'results') else {}
            opts = page_results.get("options", {})

            # Offset for this page (if the PDF was moved in the scene)
            off_x, off_y = self.page_offsets.get(page_idx, (0, 0))
            
            # Draw marks and answers
            for mark in self.view.option_marks:
                rect = mark.sceneBoundingRect()
                q_num = mark.question_num
                
                if rect:
                    x = int(rect.x() - off_x)
                    y = int(rect.y() - off_y)
                    w = int(rect.width())
                    h = int(rect.height())
                    
                    # Draw rectangle border
                    painter.setPen(QPen(QColor(0, 100, 255), 2))
                    painter.drawRect(x, y, w, h)
                    
                    # Get student answer and correct answer
                    # Ensure q_num is int for consistent key lookup
                    q_num_int = int(q_num) if isinstance(q_num, (int, str)) and str(q_num).isdigit() else q_num
                    student_answer = opts.get(q_num_int, "") or opts.get(q_num, "") or opts.get(str(q_num), "")
                    correct_answer = self.answer_key.get(q_num_int, "") or self.answer_key.get(q_num, "") or self.answer_key.get(str(q_num), "")
                    
                    # Debug print
                    if correct_answer:
                        print(f"Q{q_num}: correct={correct_answer}")
                    
                    # Calculate cell positions for A, B, C, D
                    num_options = getattr(mark, "options_count", 4)
                    cell_width = w // num_options
                    option_labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[:num_options]
                    
                    for i, opt_label in enumerate(option_labels):
                        cell_x = x + i * cell_width
                        cell_center_x = cell_x + cell_width // 2
                        cell_center_y = y + h // 2
                        
                        # Draw red dot for correct answer
                        if correct_answer and opt_label.upper() == correct_answer.upper():
                            # Must set both brush and pen before each ellipse
                            painter.save()
                            painter.setBrush(QBrush(QColor(255, 0, 0)))
                            painter.setPen(QPen(QColor(255, 0, 0), 2))
                            painter.drawEllipse(cell_center_x - 8, cell_center_y - 8, 16, 16)
                            painter.restore()
                        
                        # Draw X mark for student's wrong answer
                        if student_answer and opt_label.upper() == student_answer.upper():
                            if correct_answer and student_answer.upper() != correct_answer.upper():
                                # Wrong answer - draw X mark
                                painter.save()
                                painter.setPen(QPen(QColor(255, 0, 0), 3))
                                painter.drawLine(cell_center_x - 8, cell_center_y - 8, cell_center_x + 8, cell_center_y + 8)
                                painter.drawLine(cell_center_x + 8, cell_center_y - 8, cell_center_x - 8, cell_center_y + 8)
                                painter.restore()
                    
                    # Draw question number
                    painter.setPen(QPen(QColor(0, 0, 0), 1))
                    painter.setFont(QFont("Arial", 10, QFont.Bold))
                    painter.drawText(x - 30, y + h // 2 + 5, f"Q{q_num}")
            
            painter.end()
            
            # Save image
            output_path = os.path.join(folder, f"page_{page_idx + 1:03d}.png")
            qimg.save(output_path)
        
        progress.setValue(len(self.pdf_document))
        QMessageBox.information(self, "Done", f"Exported {len(self.pdf_document)} images to:\n{folder}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = OMRSoftware()
    window.show()
    sys.exit(app.exec_())
