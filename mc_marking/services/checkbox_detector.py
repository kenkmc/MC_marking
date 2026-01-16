"""Detect checkbox-like marks within answer cells."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import cv2
import numpy as np


@dataclass(frozen=True)
class CheckboxAnalysis:
    """Summary of the detected checkbox and its fill level."""

    fill_ratio: float
    box_area_ratio: float
    contour_area_ratio: float
    confidence: float


def analyze_checkbox(image: np.ndarray) -> Optional[CheckboxAnalysis]:
    """Return checkbox metrics for the given RGB cell image, if found."""
    if image.size == 0:
        return None

    # Normalize to 8-bit grayscale for robust thresholding
    if image.ndim == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    else:
        gray = image.astype(np.uint8)

    gray = cv2.normalize(gray, None, 0, 255, cv2.NORM_MINMAX)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, binary = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Remove small noise while preserving rectangular contours
    kernel = np.ones((3, 3), np.uint8)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)

    contours, _ = cv2.findContours(cleaned, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    cell_area = max(1.0, float(gray.shape[0] * gray.shape[1]))
    best_analysis: Optional[CheckboxAnalysis] = None
    best_score = 0.0

    for contour in contours:
        contour_area = cv2.contourArea(contour)
        if contour_area < cell_area * 0.01:
            continue

        x, y, w, h = cv2.boundingRect(contour)
        if w < 6 or h < 6:
            continue

        aspect_ratio = w / float(h)
        if not 0.6 <= aspect_ratio <= 1.4:
            continue

        approx = cv2.approxPolyDP(contour, 0.05 * cv2.arcLength(contour, True), True)
        if len(approx) < 4 or len(approx) > 6:
            continue

        box_area = float(w * h)
        box_area_ratio = box_area / cell_area
        if not 0.04 <= box_area_ratio <= 0.85:
            continue

        contour_area_ratio = contour_area / max(box_area, 1.0)
        if contour_area_ratio < 0.35:
            continue

        # Estimate fill by sampling an inner region to avoid the border
        border_margin = max(1, int(round(min(w, h) * 0.15)))
        inner_x = x + border_margin
        inner_y = y + border_margin
        inner_w = w - border_margin * 2
        inner_h = h - border_margin * 2
        if inner_w <= 2 or inner_h <= 2:
            inner_roi = cleaned[y:y + h, x:x + w]
        else:
            inner_roi = cleaned[inner_y:inner_y + inner_h, inner_x:inner_x + inner_w]

        if inner_roi.size == 0:
            continue

        fill_ratio = float(np.mean(inner_roi) / 255.0)

        # Confidence favours squareness and strong fills
        square_score = 1.0 - min(abs(aspect_ratio - 1.0), 1.0)
        coverage_score = min(1.0, contour_area_ratio * 1.5)
        fill_strength = min(1.0, fill_ratio * 1.2)
        confidence = (square_score * 0.3) + (coverage_score * 0.3) + (fill_strength * 0.4)

        score = confidence
        if score > best_score:
            best_score = score
            best_analysis = CheckboxAnalysis(
                fill_ratio=max(0.0, min(fill_ratio, 1.0)),
                box_area_ratio=box_area_ratio,
                contour_area_ratio=contour_area_ratio,
                confidence=confidence,
            )

    return best_analysis
