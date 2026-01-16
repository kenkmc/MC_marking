"""OCR utilities with support for Tesseract and PaddleOCR."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence

import cv2
import numpy as np

# Import both OCR engines
try:
    from paddleocr import PaddleOCR
    PADDLE_AVAILABLE = True
except ImportError:
    PADDLE_AVAILABLE = False
    PaddleOCR = None

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

from mc_marking.models.answer_sheet import CellResult, TableExtraction
from mc_marking.services.checkbox_detector import analyze_checkbox
from mc_marking.utils.image_utils import BoundingBox, crop


# Global PaddleOCR instance (initialized once for performance)
_paddle_ocr = None


@dataclass
class OcrConfig:
    """Parameters controlling the OCR pipeline."""

    languages: str = "eng"
    psm_mode: int = 6  # Assume a uniform block of text
    engine: str = "paddle"  # "paddle" or "tesseract"


def _get_paddle_ocr():
    """Get or initialize the global PaddleOCR instance."""
    global _paddle_ocr
    if _paddle_ocr is None and PADDLE_AVAILABLE:
        # Initialize with only valid parameters
        _paddle_ocr = PaddleOCR(
            use_angle_cls=True,  # Enable text orientation detection
            lang='en',           # English language
        )
    return _paddle_ocr


def _coerce_region_boxes(raw: Optional[Sequence[object]]) -> List[BoundingBox]:
    boxes: List[BoundingBox] = []
    if not raw:
        return boxes
    for entry in raw:
        if isinstance(entry, BoundingBox):
            boxes.append(entry)
            continue
        if isinstance(entry, dict):
            try:
                boxes.append(
                    BoundingBox(
                        x=int(round(entry.get("x", 0))),
                        y=int(round(entry.get("y", 0))),
                        width=int(round(entry.get("width", 0))),
                        height=int(round(entry.get("height", 0))),
                    )
                )
                continue
            except (TypeError, ValueError):
                continue
    return boxes


def _boxes_intersect(a: BoundingBox, b: BoundingBox) -> bool:
    return not (
        a.x + a.width <= b.x
        or b.x + b.width <= a.x
        or a.y + a.height <= b.y
        or b.y + b.height <= a.y
    )


def _is_in_regions(box: BoundingBox, regions: Sequence[BoundingBox]) -> bool:
    for region in regions:
        if _boxes_intersect(box, region):
            return True
    return False


def recognise_table_cells(
    image: np.ndarray,
    extraction: TableExtraction,
    config: OcrConfig | None = None,
    *,
    region_config: Optional[Dict[str, Sequence[object]]] = None,
) -> List[CellResult]:
    """Apply OCR to each cell in the detected table."""
    cfg = config or OcrConfig()
    results: List[CellResult] = []

    ocr_regions: List[BoundingBox] = []
    omr_regions: List[BoundingBox] = []
    if region_config:
        ocr_regions = _coerce_region_boxes(region_config.get("ocr"))
        omr_regions = _coerce_region_boxes(region_config.get("omr"))
    
    total_cells = len(extraction.cells)
    engine_name = cfg.engine.upper()
    print(f"OCR: Using {engine_name} engine to process {total_cells} cells from table with {extraction.row_count} rows x {extraction.column_count} cols")
    
    # Check if selected engine is available
    if cfg.engine == "paddle" and not PADDLE_AVAILABLE:
        print(f"  Warning: PaddleOCR not available, falling back to Tesseract")
        cfg.engine = "tesseract"
    
    if cfg.engine == "tesseract" and not TESSERACT_AVAILABLE:
        print(f"  Error: No OCR engine available! Please install paddleocr or pytesseract")
        return results
    
    for idx, cell in enumerate(extraction.cells):
        # Progress indicator every 50 cells
        if idx > 0 and idx % 50 == 0:
            print(f"  Progress: {idx}/{total_cells} cells processed...")

        cell_box = cell.bounding_box
        in_ocr_region = _is_in_regions(cell_box, ocr_regions) if ocr_regions else False
        in_omr_region = _is_in_regions(cell_box, omr_regions) if omr_regions else False
        treat_as_omr = in_omr_region and not in_ocr_region
        perform_ocr = not treat_as_omr

        cell_image = crop(image, cell_box)

        # Skip empty or tiny cells
        if cell_image.shape[0] < 5 or cell_image.shape[1] < 5:
            results.append(
                CellResult(
                    row=cell.row,
                    column=cell.column,
                    text="",
                    confidence=0.0,
                    bounding_box=cell_box,
                    ink_density=0.0,
                )
            )
            continue

        checkbox_analysis = None
        if cell.column > 0 and cell.row > 0:
            checkbox_analysis = analyze_checkbox(cell_image)

        text = ""
        confidence = 0.0

        if perform_ocr:
            padded_cell = _add_padding(cell_image, padding=8)
            ocr_ready, _, _ = _preprocess_cell_for_ocr(padded_cell)

            if cfg.engine == "paddle":
                text, confidence = _run_paddle_ocr(ocr_ready, cell.column)
                if not text and TESSERACT_AVAILABLE:
                    fallback_text, fallback_confidence = _run_tesseract_ocr(ocr_ready, cell.column, cfg)
                    if fallback_text:
                        text = fallback_text
                        confidence = fallback_confidence
            else:
                text, confidence = _run_tesseract_ocr(ocr_ready, cell.column, cfg)

            if cell.row == 0:
                print(f"  Row 0, Col {cell.column}: '{text}' (confidence: {confidence:.2f})")
        else:
            if checkbox_analysis is not None:
                confidence = checkbox_analysis.confidence
            else:
                confidence = _estimate_confidence(cell_image)

        base_ink_density = _estimate_ink_density(cell_image)
        if checkbox_analysis is not None:
            ink_density = max(base_ink_density, checkbox_analysis.fill_ratio)
        else:
            ink_density = base_ink_density
        
        results.append(
            CellResult(
                row=cell.row,
                column=cell.column,
                text=text,
                confidence=confidence,
                bounding_box=cell_box,
                ink_density=ink_density,
            )
        )
    
    print(f"OCR: Completed. Found {sum(1 for r in results if r.text)} cells with text")
    return results


def _run_paddle_ocr(image: np.ndarray, column_index: int) -> tuple[str, float]:
    """Run PaddleOCR on a cell image."""
    if not PADDLE_AVAILABLE:
        return "", 0.0
    
    try:
        ocr = _get_paddle_ocr()
        if ocr is None:
            return "", 0.0

        # PaddleOCR expects RGB images; convert if needed
        if len(image.shape) == 2 or image.shape[2] == 1:
            sample = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
        else:
            sample = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

        # PaddleOCR enforces a maximum side length (default 4000); shrink over-sized cells to comply
        max_side = max(sample.shape[0], sample.shape[1])
        if max_side > 4000:
            scale = 4000.0 / max_side
            new_size = (max(1, int(sample.shape[1] * scale)), max(1, int(sample.shape[0] * scale)))
            sample = cv2.resize(sample, new_size, interpolation=cv2.INTER_AREA)

        # Run OCR on the prepared sample
        result = ocr.ocr(sample)

        if not result:
            return "", 0.0

        best_text = ""
        best_conf = 0.0

        def _collect_entries(node) -> List[tuple[str, float]]:
            collected: List[tuple[str, float]] = []
            if isinstance(node, str):
                collected.append((node, 0.0))
            elif isinstance(node, (list, tuple)):
                if len(node) == 2 and isinstance(node[0], str):
                    score = float(node[1]) if isinstance(node[1], (int, float)) else 0.0
                    collected.append((node[0], score))
                else:
                    for item in node:
                        collected.extend(_collect_entries(item))
            return collected

        candidates: List[tuple[str, float]] = []
        for entry in result:
            candidates.extend(_collect_entries(entry))

        for candidate_text, candidate_conf in candidates:
            if candidate_text:
                if candidate_conf > best_conf or not best_text:
                    best_text = str(candidate_text)
                    best_conf = float(candidate_conf)
        
        cleaned = _clean_cell_text(best_text)
        
        # Validate based on column type
        if column_index == 0:
            # Question numbers - should be digits
            if cleaned and not cleaned.isdigit():
                # Try to extract digits only
                digits_only = "".join(c for c in cleaned if c.isdigit())
                if digits_only:
                    cleaned = digits_only
        else:
            # Choice letters - should be single letter
            if cleaned and not cleaned.isalpha():
                # Try to extract letters only
                letters_only = "".join(c for c in cleaned if c.isalpha())
                if letters_only:
                    cleaned = letters_only[:1]  # Take first letter only
        
        return cleaned, float(best_conf)
        
    except Exception as e:
        print(f"  PaddleOCR error: {e}")
        return "", 0.0


def _run_tesseract_ocr(image: np.ndarray, column_index: int, cfg: OcrConfig) -> tuple[str, float]:
    """Run Tesseract OCR on a cell image."""
    if not TESSERACT_AVAILABLE:
        return "", 0.0
    
    text = ""
    
    # Strategy 1: Use appropriate PSM mode
    if column_index == 0:
        # Question numbers - use line mode WITHOUT whitelist first
        config_str = "--psm 7 --oem 3"
    else:
        # Choice letters - try without whitelist too
        config_str = "--psm 10 --oem 3"
    
    try:
        text = pytesseract.image_to_string(image, lang=cfg.languages, config=config_str)
        cleaned = _clean_cell_text(text)
        
        # If we got something but it's not what we expect, try with whitelist
        if column_index == 0 and cleaned and not cleaned.isdigit():
            # Retry with digit whitelist
            text2 = pytesseract.image_to_string(image, lang=cfg.languages, 
                                               config="--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789")
            cleaned2 = _clean_cell_text(text2)
            if cleaned2 and cleaned2.isdigit():
                cleaned = cleaned2
        elif column_index > 0 and cleaned and not cleaned.isalpha():
            # Retry with letter whitelist
            text2 = pytesseract.image_to_string(image, lang=cfg.languages,
                                               config="--psm 10 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ")
            cleaned2 = _clean_cell_text(text2)
            if cleaned2 and cleaned2.isalpha():
                cleaned = cleaned2
                
    except Exception as e:
        print(f"OCR error: {e}")
        cleaned = ""
    
    # Strategy 2: If failed, try fallback methods
    if not cleaned:
        cleaned = _run_fallback_tesseract_passes(image, column_index)
    
    confidence = _estimate_confidence(image)
    return cleaned, confidence


def _estimate_confidence(image: np.ndarray) -> float:
    variance = float(np.var(image))
    normalized = min(1.0, max(0.0, variance / (255.0 ** 2)))
    return normalized


def _estimate_ink_density(image: np.ndarray, *, precomputed_mask: np.ndarray | None = None) -> float:
    if precomputed_mask is not None:
        return float(np.mean(precomputed_mask / 255.0))
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    return float(np.mean(binary / 255.0))


def _add_padding(image: np.ndarray, padding: int = 5) -> np.ndarray:
    """Add white padding around the image for better OCR boundary detection."""
    return cv2.copyMakeBorder(
        image,
        padding,
        padding,
        padding,
        padding,
        cv2.BORDER_CONSTANT,
        value=(255, 255, 255)
    )


def _preprocess_cell_for_ocr(cell_image: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    # Handle very small images
    if cell_image.shape[0] < 10 or cell_image.shape[1] < 10:
        return cell_image, cell_image, cell_image
    
    gray = cv2.cvtColor(cell_image, cv2.COLOR_RGB2GRAY)
    
    # Upscale small images first for better processing
    h, w = gray.shape
    scale_factor = 1
    if min(h, w) < 20:
        scale_factor = 4
    elif min(h, w) < 30:
        scale_factor = 3
    elif min(h, w) < 50:
        scale_factor = 2
    
    if scale_factor > 1:
        new_w = w * scale_factor
        new_h = h * scale_factor
        gray = cv2.resize(gray, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    
    # Simple bilateral filter instead of heavy denoising
    denoised = cv2.bilateralFilter(gray, 5, 50, 50)
    
    # Enhance contrast
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(4, 4))
    enhanced = clahe.apply(denoised)
    
    # Simple Otsu thresholding
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Minimal morphological cleanup
    kernel = np.ones((2, 2), np.uint8)
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel, iterations=1)
    
    # Light sharpen
    kernel_sharpen = np.array([[0, -1, 0],
                               [-1, 5, -1],
                               [0, -1, 0]])
    binary = cv2.filter2D(binary, -1, kernel_sharpen)
    
    # Final threshold
    _, binary = cv2.threshold(binary, 127, 255, cv2.THRESH_BINARY)
    
    mask = 255 - binary
    return binary, mask, enhanced


def _run_fallback_ocr_passes(image: np.ndarray, column_index: int) -> str:
    """Fallback OCR passes for Tesseract - renamed for backwards compatibility."""
    return _run_fallback_tesseract_passes(image, column_index)


def _run_fallback_tesseract_passes(image: np.ndarray, column_index: int) -> str:
    attempts = []
    if column_index == 0:
        # Question numbers - try different PSM modes with digit whitelist
        attempts.extend([
            "--psm 8 --oem 3 -c tessedit_char_whitelist=0123456789",
            "--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789",
            "--psm 13 --oem 3 -c tessedit_char_whitelist=0123456789",  # Raw line mode
            "--psm 7 --oem 3",  # Without whitelist
        ])
    else:
        # Choice letters - try different modes
        attempts.extend([
            "--psm 8 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ",
            "--psm 13 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ",
            "--psm 10 --oem 3",  # Without whitelist
            "--psm 8 --oem 3",  # Without whitelist
        ])
    
    for config in attempts:
        try:
            candidate = pytesseract.image_to_string(image, config=config)
            cleaned = _clean_cell_text(candidate)
            if cleaned:
                return cleaned
        except Exception:
            continue
    
    return ""


def _clean_cell_text(text: str) -> str:
    """Clean and normalize OCR output for MC sheet cells."""
    if not text:
        return ""
    
    # Remove whitespace and newlines
    normalized = text.replace("\n", " ").replace("\r", " ").replace("\t", " ").strip()
    
    # Common OCR corrections
    corrections = {
        "|": "1",
        "!": "1", 
        "l": "1",
        "I": "1",  # Capital I to 1 for numbers
        "O": "0",
        "o": "0",
        "S": "5",
        "s": "5",
        "Z": "2",
        "z": "2",
        "B": "8",
        "G": "6",
        "D": "0",
    }
    
    # Apply corrections
    result = ""
    for char in normalized:
        result += corrections.get(char, char)
    
    # Remove non-alphanumeric
    cleaned = "".join(c for c in result if c.isalnum())
    
    if not cleaned:
        return ""
    
    # Return uppercase
    return cleaned.upper()
