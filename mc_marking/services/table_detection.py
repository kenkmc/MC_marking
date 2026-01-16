"""Detect and segment answer tables from scanned pages."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

import cv2
import numpy as np

from mc_marking.models.answer_sheet import CellResult, TableExtraction
from mc_marking.utils.image_utils import BoundingBox, crop


@dataclass
class TableDetectionConfig:
    """Configure morphological operations for table detection."""

    min_table_area_ratio: float = 0.08  # Reduced to catch smaller table sections
    kernel_scale: float = 0.008  # Reduced to preserve finer details
    min_cell_size: int = 8  # Smaller to detect individual answer bubbles
    merge_nearby_tables: bool = True  # Merge tables that are part of the same sheet
    max_table_gap: int = 50  # Maximum horizontal gap between tables to merge


def detect_table(
    image: np.ndarray,
    page_index: int,
    source_path: str,
    config: TableDetectionConfig | None = None,
    roi: BoundingBox | None = None,
) -> Optional[TableExtraction]:
    """Return the most prominent table detected within the provided region."""

    tables = detect_tables(image, page_index, source_path, config=config, roi=roi)
    return tables[0] if tables else None


def detect_tables(
    image: np.ndarray,
    page_index: int,
    source_path: str,
    config: TableDetectionConfig | None = None,
    roi: BoundingBox | None = None,
) -> List[TableExtraction]:
    """Detect all answer tables visible within the image or ROI."""

    cfg = config or TableDetectionConfig()
    search_image = crop(image, roi) if roi else image
    offset_x = roi.x if roi else 0
    offset_y = roi.y if roi else 0

    prepared = _prepare_binary_mask(search_image)
    kernel_size = max(3, int(cfg.kernel_scale * max(search_image.shape[:2])))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
    morph = cv2.morphologyEx(prepared, cv2.MORPH_CLOSE, kernel, iterations=2)

    contours, _ = cv2.findContours(morph, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return []

    h, w = search_image.shape[:2]
    min_area = cfg.min_table_area_ratio * h * w
    candidates = [cnt for cnt in contours if cv2.contourArea(cnt) >= min_area]
    
    # Get bounding boxes for all candidates
    bboxes = []
    for contour in candidates:
        x, y, width, height = cv2.boundingRect(contour)
        bboxes.append(BoundingBox(x=x + offset_x, y=y + offset_y, width=width, height=height))
    
    # Merge nearby tables if enabled
    if cfg.merge_nearby_tables and len(bboxes) > 1:
        bboxes = _merge_nearby_tables(bboxes, cfg.max_table_gap)
    
    tables: List[TableExtraction] = []
    for bounding_box in bboxes:
        extraction = _extract_table(
            image=image,
            bounding_box=bounding_box,
            page_index=page_index,
            source_path=source_path,
            config=cfg,
        )
        if extraction is not None and extraction.cells:
            tables.append(extraction)

    tables.sort(key=lambda table: (table.bounding_box.y, table.bounding_box.x))
    return tables


def _prepare_binary_mask(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    return cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 31, 12)


def _merge_nearby_tables(bboxes: List[BoundingBox], max_gap: int) -> List[BoundingBox]:
    """Merge tables that are horizontally or vertically aligned and close together."""
    if not bboxes:
        return []
    
    # Sort by y-position first, then x-position
    sorted_boxes = sorted(bboxes, key=lambda b: (b.y, b.x))
    merged = []
    
    # Group boxes by vertical position (same row)
    rows: List[List[BoundingBox]] = []
    current_row = [sorted_boxes[0]]
    
    for box in sorted_boxes[1:]:
        # Check if this box is in the same horizontal band as current row
        prev_box = current_row[-1]
        y_overlap = (min(prev_box.y + prev_box.height, box.y + box.height) - 
                     max(prev_box.y, box.y))
        
        if y_overlap > 0 or abs(box.y - current_row[0].y) < max_gap:
            current_row.append(box)
        else:
            rows.append(current_row)
            current_row = [box]
    
    if current_row:
        rows.append(current_row)
    
    # Merge horizontally adjacent boxes in each row
    for row in rows:
        if not row:
            continue
        
        # Sort by x-position
        row_sorted = sorted(row, key=lambda b: b.x)
        
        # Merge boxes that are close horizontally
        current_merge = row_sorted[0]
        
        for box in row_sorted[1:]:
            # Check horizontal gap
            gap = box.x - (current_merge.x + current_merge.width)
            
            if gap <= max_gap:
                # Merge boxes
                new_x = min(current_merge.x, box.x)
                new_y = min(current_merge.y, box.y)
                new_right = max(current_merge.x + current_merge.width, box.x + box.width)
                new_bottom = max(current_merge.y + current_merge.height, box.y + box.height)
                
                current_merge = BoundingBox(
                    x=new_x,
                    y=new_y,
                    width=new_right - new_x,
                    height=new_bottom - new_y
                )
            else:
                # Gap too large, save current merge and start new one
                merged.append(current_merge)
                current_merge = box
        
        merged.append(current_merge)
    
    return merged


def _extract_table(
    *,
    image: np.ndarray,
    bounding_box: BoundingBox,
    page_index: int,
    source_path: str,
    config: TableDetectionConfig,
) -> Optional[TableExtraction]:
    table_roi = crop(image, bounding_box)
    row_positions, col_positions = _estimate_grid(table_roi, config)
    if row_positions is None or col_positions is None:
        return None

    cells: List[CellResult] = []
    for r_idx, (r_start, r_end) in enumerate(_pairwise(row_positions)):
        for c_idx, (c_start, c_end) in enumerate(_pairwise(col_positions)):
            cell_crop = table_roi[r_start:r_end, c_start:c_end]
            if cell_crop.size == 0:
                continue
            average_intensity = float(np.mean(cell_crop))
            cell_box = BoundingBox(
                x=bounding_box.x + c_start,
                y=bounding_box.y + r_start,
                width=c_end - c_start,
                height=r_end - r_start,
            )
            cells.append(
                CellResult(
                    row=r_idx,
                    column=c_idx,
                    text="",
                    confidence=average_intensity / 255.0,
                    bounding_box=cell_box,
                    ink_density=0.0,
                )
            )

    return TableExtraction(
        source_path=Path(source_path),
        page_index=page_index,
        bounding_box=bounding_box,
        cells=cells,
        row_count=len(row_positions) - 1,
        column_count=len(col_positions) - 1,
    )


def _estimate_grid(image: np.ndarray, config: TableDetectionConfig) -> Tuple[Optional[List[int]], Optional[List[int]]]:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    _, projected_bin = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Use smaller kernels to detect finer grid lines
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(3, int(image.shape[0] * 0.03))))
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(3, int(image.shape[1] * 0.03)), 1))
    vertical_lines = cv2.morphologyEx(projected_bin, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    horizontal_lines = cv2.morphologyEx(projected_bin, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    grid_mask = cv2.add(vertical_lines, horizontal_lines)

    # More sensitive edge detection for fine lines
    edges = cv2.Canny(grid_mask, 30, 100, apertureSize=3)
    
    # Detect lines with lower threshold for finer grids
    lines = cv2.HoughLinesP(
        edges,
        1,
        np.pi / 180,
        threshold=max(50, int(min(image.shape[:2]) * 0.1)),  # Adaptive threshold
        minLineLength=config.min_cell_size * 2,  # Shorter lines acceptable
        maxLineGap=config.min_cell_size * 2,  # Larger gap tolerance
    )

    rows: List[int] = [0, image.shape[0]]
    cols: List[int] = [0, image.shape[1]]
    
    if lines is not None:
        for x1, y1, x2, y2 in lines[:, 0]:
            # Vertical lines (columns)
            if abs(x1 - x2) < 15:  # Allow slightly diagonal lines
                cols.append(min(x1, x2))
            # Horizontal lines (rows)
            if abs(y1 - y2) < 15:  # Allow slightly diagonal lines
                rows.append(min(y1, y2))

    # If we didn't find enough lines, try projection method
    if len(rows) < 3 or len(cols) < 3:
        row_proj, col_proj = _estimate_grid_by_projection(gray, config)
        if row_proj and len(row_proj) > len(rows):
            rows = row_proj
        if col_proj and len(col_proj) > len(cols):
            cols = col_proj

    rows = _filter_positions(sorted(set(rows)), config.min_cell_size)
    cols = _filter_positions(sorted(set(cols)), config.min_cell_size)

    if len(rows) < 2 or len(cols) < 2:
        return None, None
    return rows, cols


def _pairwise(positions: Sequence[int]) -> List[Tuple[int, int]]:
    return list(zip(positions[:-1], positions[1:]))


def _estimate_grid_by_projection(image: np.ndarray, config: TableDetectionConfig) -> Tuple[List[int], List[int]]:
    """Estimate grid lines using projection profile method as fallback."""
    # Horizontal projection to find rows
    horizontal_proj = np.sum(image, axis=1)
    h_mean = np.mean(horizontal_proj)
    h_threshold = h_mean * 0.7
    
    # Find transitions in projection (potential row boundaries)
    rows = [0]
    in_content = False
    min_row_height = config.min_cell_size
    
    for i in range(1, len(horizontal_proj)):
        if not in_content and horizontal_proj[i] > h_threshold:
            if i - rows[-1] >= min_row_height:
                rows.append(i)
            in_content = True
        elif in_content and horizontal_proj[i] < h_threshold:
            if i - rows[-1] >= min_row_height:
                rows.append(i)
            in_content = False
    
    rows.append(image.shape[0])
    
    # Vertical projection to find columns
    vertical_proj = np.sum(image, axis=0)
    v_mean = np.mean(vertical_proj)
    v_threshold = v_mean * 0.7
    
    cols = [0]
    in_content = False
    min_col_width = config.min_cell_size
    
    for i in range(1, len(vertical_proj)):
        if not in_content and vertical_proj[i] > v_threshold:
            if i - cols[-1] >= min_col_width:
                cols.append(i)
            in_content = True
        elif in_content and vertical_proj[i] < v_threshold:
            if i - cols[-1] >= min_col_width:
                cols.append(i)
            in_content = False
    
    cols.append(image.shape[1])
    
    return rows, cols


def _filter_positions(positions: Sequence[int], min_distance: int) -> List[int]:
    ordered = list(positions)
    if not ordered:
        return []
    filtered: List[int] = [ordered[0]]
    for pos in ordered[1:-1]:
        if abs(pos - filtered[-1]) >= min_distance:
            filtered.append(pos)
    if abs(ordered[-1] - filtered[-1]) >= min_distance:
        filtered.append(ordered[-1])
    elif ordered[-1] != filtered[-1]:
        filtered[-1] = ordered[-1]
    return filtered
