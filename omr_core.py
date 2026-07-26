"""Fast, testable image-processing primitives used by CheckMate.

This module deliberately has no Qt dependencies.  Keeping the OMR decision
logic here makes it possible to test recognition without starting the desktop
application and lets the GUI avoid repeated PIL/NumPy conversions.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Mapping, Sequence

import cv2
import numpy as np


OPTION_LABELS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@dataclass(frozen=True)
class OMRDetectionResult:
    """Structured result for one multiple-choice answer region."""

    answer: str
    confidence: float
    needs_review: bool
    reason: str
    cell_scores: tuple[dict[str, Any], ...]
    content_bounds: tuple[int, int]
    cell_edges: tuple[int, ...]
    decision: dict[str, Any]

    def as_record(self) -> dict[str, Any]:
        """Return a JSON-serialisable representation."""

        return asdict(self)


def _as_rgb_uint8(image: Any) -> np.ndarray:
    """Convert a PIL image or array-like object to contiguous uint8 RGB."""

    array = np.asarray(image)
    if array.size == 0:
        raise ValueError("The answer region is empty.")

    if array.dtype != np.uint8:
        if np.issubdtype(array.dtype, np.floating) and float(np.nanmax(array)) <= 1.0:
            array = array * 255.0
        array = np.nan_to_num(array, nan=255.0, posinf=255.0, neginf=0.0)
        array = np.clip(array, 0, 255).astype(np.uint8)

    if array.ndim == 2:
        return cv2.cvtColor(np.ascontiguousarray(array), cv2.COLOR_GRAY2RGB)
    if array.ndim != 3:
        raise ValueError(f"Unsupported answer-region shape: {array.shape!r}")
    if array.shape[2] == 1:
        return cv2.cvtColor(np.ascontiguousarray(array[:, :, 0]), cv2.COLOR_GRAY2RGB)
    if array.shape[2] >= 4:
        # PIL and the rest of CheckMate use RGBA ordering.
        rgb = array[:, :, :3].astype(np.float32)
        alpha = array[:, :, 3:4].astype(np.float32) / 255.0
        array = rgb * alpha + 255.0 * (1.0 - alpha)
    return np.ascontiguousarray(array[:, :, :3].astype(np.uint8))


def estimate_option_content_bounds(
    gray: np.ndarray,
    saturation: np.ndarray,
    dark_pixel_threshold: float,
    color_pixel_threshold: float,
    options_count: int,
) -> tuple[int, int]:
    """Estimate the horizontal span containing the option boxes.

    The fallback to the full width is intentional.  A narrow activity span can
    simply mean that only one option was marked, so aggressively trimming it
    would shift the cell boundaries and create a false answer.
    """

    height, width = gray.shape[:2]
    activity_mask = np.logical_or(
        gray < dark_pixel_threshold,
        saturation > color_pixel_threshold,
    )
    col_activity = np.mean(activity_mask, axis=0)
    min_col_activity = max(0.08, 2.0 / max(1, height))
    active_cols = np.flatnonzero(col_activity >= min_col_activity)
    if active_cols.size == 0:
        return 0, width

    left = max(0, int(active_cols[0]) - 1)
    right = min(width, int(active_cols[-1]) + 2)
    min_span = max(options_count * 5, int(round(width * 0.62)))
    if right - left < min_span:
        return 0, width
    return left, right


def _safe_inner(array: np.ndarray, x_ratio: float, y_ratio: float) -> np.ndarray:
    height, width = array.shape[:2]
    margin_x = min(max(1, int(round(width * x_ratio))), max(0, width // 2 - 1))
    margin_y = min(max(1, int(round(height * y_ratio))), max(0, height // 2 - 1))
    inner = array[margin_y : height - margin_y, margin_x : width - margin_x]
    return inner if inner.size else array


def _estimate_repeated_rectangle_geometry(
    gray: np.ndarray,
    cell_edges: Sequence[int],
) -> tuple[tuple[float, float, float, float], ...] | None:
    """Locate a repeated row of hollow rectangular response boxes.

    Some legacy templates tightly cover the four response rectangles but leave
    them close to the crop's top edge.  Measuring the geometric crop centre in
    those templates samples paper below the boxes and lets chromatic print
    fringes dominate.  A response rectangle is wider than it is tall and is
    repeated at the same relative position in every option cell, so a median
    consensus can relocate the masks without following one student's mark.

    Returns one local ``(center_x, center_y, width, height)`` tuple per cell.
    Circular bubbles and rows without a strong three-cell consensus
    intentionally fall back to the historical crop-centred geometry.
    """

    height = int(gray.shape[0])
    candidates: list[tuple[int, float, float, float, float]] = []
    cell_widths: list[int] = []
    for cell_index, (left, right) in enumerate(zip(cell_edges, cell_edges[1:])):
        cell = gray[:, left:right]
        cell_width = int(cell.shape[1])
        cell_widths.append(cell_width)
        if cell_width < 8 or height < 6:
            continue

        paper = float(np.percentile(cell, 82))
        # A slightly wider threshold retains the complete antialiased outer
        # box instead of returning a smaller, shifted fragment.  The repeated
        # four-cell consensus rejects paper texture and handwriting.
        component_mask = (cell.astype(np.float32) <= paper - 18.0).astype(np.uint8)
        contours, _ = cv2.findContours(
            component_mask,
            cv2.RETR_LIST,
            cv2.CHAIN_APPROX_SIMPLE,
        )
        located: list[tuple[float, int, int, int, int]] = []
        for contour in contours:
            x, y, width, component_height = cv2.boundingRect(contour)
            aspect = width / max(1.0, float(component_height))
            if (
                aspect <= 1.4
                or width < max(9, int(round(cell_width * 0.30)))
                or width > cell_width
                or component_height < 4
                or component_height > max(12, int(round(height * 0.75)))
            ):
                continue
            contour_area = float(cv2.contourArea(contour))
            score = (
                width * 4.0
                + contour_area * 0.08
                - max(0.0, y - height * 0.45) * 3.0
            )
            located.append((score, x, y, width, component_height))

        if not located:
            continue
        _, x, y, width, component_height = max(
            located,
            key=lambda item: item[0],
        )
        candidates.append(
            (
                cell_index,
                (x + (width - 1.0) / 2.0) / cell_width,
                y + (component_height - 1.0) / 2.0,
                width / cell_width,
                float(component_height),
            )
        )

    # Requiring three agreeing cells prevents one or two handwritten strokes
    # from redefining fixed option geometry.
    if len(candidates) < 3:
        return None

    median = tuple(
        float(np.median([row[index] for row in candidates]))
        for index in range(1, 5)
    )
    inliers = [
        row
        for row in candidates
        if abs(row[1] - median[0]) <= 0.18
        and abs(row[2] - median[1]) <= max(2.5, height * 0.14)
        and abs(row[3] - median[2]) <= 0.24
        and abs(row[4] - median[3]) <= max(2.5, height * 0.14)
    ]
    if len(inliers) < 3:
        return None
    median = tuple(
        float(np.median([row[index] for row in inliers]))
        for index in range(1, 5)
    )
    by_cell = {int(row[0]): row[1:] for row in inliers}
    global_centers = [
        (
            int(row[0]),
            cell_edges[int(row[0])]
            + float(row[1]) * cell_widths[int(row[0])],
        )
        for row in inliers
    ]
    spacing_candidates = [
        (right_center - left_center) / (right_index - left_index)
        for left_position, (left_index, left_center) in enumerate(global_centers)
        for right_index, right_center in global_centers[left_position + 1 :]
        if right_index != left_index
    ]
    fitted_spacing = float(np.median(spacing_candidates))
    fitted_origin = float(
        np.median(
            [
                center - cell_index * fitted_spacing
                for cell_index, center in global_centers
            ]
        )
    )
    median_width = float(
        np.median(
            [
                float(row[3]) * cell_widths[int(row[0])]
                for row in inliers
            ]
        )
    )
    geometry = []
    for cell_index, cell_width in enumerate(cell_widths):
        located = by_cell.get(cell_index)
        if located is None:
            center_x = (
                fitted_origin
                + cell_index * fitted_spacing
                - cell_edges[cell_index]
            )
            center_y = median[1]
            target_width = median_width
            component_height = median[3]
        else:
            relative_x, center_y, relative_width, component_height = located
            center_x = relative_x * cell_width
            target_width = relative_width * cell_width
        geometry.append(
            (
                _clamp(center_x, 0.0, cell_width - 1.0),
                _clamp(center_y, 0.0, height - 1.0),
                max(4.0, target_width),
                max(3.0, component_height),
            )
        )
    return tuple(geometry)


def _mean(values: Sequence[float]) -> float:
    return float(sum(values) / max(1, len(values)))


def _clamp(value: float, minimum: float = 0.0, maximum: float = 1.0) -> float:
    return max(minimum, min(maximum, float(value)))


def _select_interior_scores(
    raw_scores: Sequence[Mapping[str, Any]],
) -> tuple[list[str], dict[str, Any]]:
    """Select marks from outline-resistant interior features.

    Printed bubbles and table borders are mostly thin, hollow contours.  A
    real response puts ink through the centre of the bubble and normally
    creates strokes that are thicker than the printed outline.  These gates
    intentionally use absolute, per-cell evidence as the primary decision;
    the blank-cell baseline is only a sensitivity aid.  Consequently adding
    a darker second mark cannot raise the threshold and hide an existing one.
    """

    scores = [dict(score) for score in raw_scores]
    evidence_values = [float(score["interior_evidence"]) for score in scores]
    baseline = min(evidence_values)

    feature_names = (
        "interior_evidence",
        "core_dark_signal",
        "core_ink_ratio",
        "nucleus_dark_signal",
        "nucleus_ink_ratio",
        "core_thick_ratio",
        "core_color_signal",
        "core_color_ratio",
        "nucleus_color_signal",
        "nucleus_color_ratio",
        "nucleus_local_signal",
        "nucleus_peak_signal",
    )
    baselines = {
        feature: min(float(score.get(feature, 0.0)) for score in scores)
        for feature in feature_names
    }
    for score in scores:
        for feature in feature_names:
            score[f"{feature}_margin"] = (
                float(score.get(feature, 0.0)) - baselines[feature]
            )

    thresholds = {
        "min_evidence": 4.8,
        "min_relative_evidence": 3.2,
        "min_core_signal": 5.0,
        "min_core_ratio": 0.055,
        "min_nucleus_signal": 4.5,
        "min_nucleus_ratio": 0.045,
        "min_thick_ratio": 0.012,
        "min_color_ratio": 0.035,
        "min_nucleus_local_signal": 12.0,
        "min_nucleus_peak_signal": 20.0,
        "rectangle_color_core_ratio": 0.20,
        "rectangle_color_nucleus_ratio": 0.18,
        "rectangle_color_signal": 15.0,
        "rectangle_color_dark_signal": 18.0,
        "rectangle_dark_core_signal": 22.0,
        "rectangle_dark_nucleus_signal": 20.0,
        "rectangle_dark_peak_signal": 25.0,
        "rectangle_dark_core_ratio": 0.30,
        "rectangle_dark_nucleus_ratio": 0.25,
        "rectangle_relative_evidence_margin": 10.0,
        "rectangle_relative_core_margin": 7.0,
        "rectangle_relative_nucleus_margin": 6.0,
        "rectangle_relative_peak_margin": 10.0,
    }

    candidates: list[dict[str, Any]] = []
    for score in scores:
        evidence = float(score["interior_evidence"])
        evidence_margin = evidence - baseline
        core_signal = float(score.get("core_dark_signal", 0.0))
        core_ratio = float(score.get("core_ink_ratio", 0.0))
        nucleus_signal = float(score.get("nucleus_dark_signal", 0.0))
        nucleus_ratio = float(score.get("nucleus_ink_ratio", 0.0))
        thick_ratio = float(score.get("core_thick_ratio", 0.0))
        color_signal = float(score.get("core_color_signal", 0.0))
        color_ratio = float(score.get("core_color_ratio", 0.0))
        nucleus_color_ratio = float(score.get("nucleus_color_ratio", 0.0))
        nucleus_local_signal = float(score.get("nucleus_local_signal", 0.0))
        nucleus_peak_signal = float(score.get("nucleus_peak_signal", 0.0))

        if score.get("target_source") == "repeated_rectangle":
            rectangle_color_mark = (
                color_ratio >= thresholds["rectangle_color_core_ratio"]
                and nucleus_color_ratio
                >= thresholds["rectangle_color_nucleus_ratio"]
                and color_signal >= thresholds["rectangle_color_signal"]
                and core_signal >= thresholds["rectangle_color_dark_signal"]
                and nucleus_signal
                >= thresholds["rectangle_color_dark_signal"]
            )
            rectangle_dark_mark = (
                core_signal >= thresholds["rectangle_dark_core_signal"]
                and nucleus_signal
                >= thresholds["rectangle_dark_nucleus_signal"]
                and nucleus_peak_signal
                >= thresholds["rectangle_dark_peak_signal"]
                and core_ratio >= thresholds["rectangle_dark_core_ratio"]
                and nucleus_ratio
                >= thresholds["rectangle_dark_nucleus_ratio"]
            )
            rectangle_relative_mark = (
                float(score.get("interior_evidence_margin", 0.0))
                >= thresholds["rectangle_relative_evidence_margin"]
                and float(score.get("core_dark_signal_margin", 0.0))
                >= thresholds["rectangle_relative_core_margin"]
                and float(score.get("nucleus_dark_signal_margin", 0.0))
                >= thresholds["rectangle_relative_nucleus_margin"]
                and float(score.get("nucleus_peak_signal_margin", 0.0))
                >= thresholds["rectangle_relative_peak_margin"]
                and core_signal >= 12.0
                and core_ratio >= 0.25
                and nucleus_ratio >= 0.15
            )
            score["strength"] = float(
                evidence / max(thresholds["min_evidence"], 1e-6)
                + core_signal
                / max(thresholds["rectangle_dark_core_signal"], 1e-6)
                + nucleus_signal
                / max(thresholds["rectangle_dark_nucleus_signal"], 1e-6)
            ) / 3.0
            score["gate"] = (
                "rectangle_color"
                if rectangle_color_mark
                else "rectangle_dark"
                if rectangle_dark_mark
                else "rectangle_relative"
                if rectangle_relative_mark
                else "none"
            )
            if (
                rectangle_color_mark
                or rectangle_dark_mark
                or rectangle_relative_mark
            ):
                candidates.append(score)
            continue

        # Independent gates recognise all-filled rows and guarantee monotonic
        # multi-select behaviour.  Requiring nucleus evidence is what rejects
        # a thick but hollow printed outline.
        dark_absolute = (
            nucleus_signal >= thresholds["min_nucleus_signal"]
            and nucleus_ratio >= thresholds["min_nucleus_ratio"]
            and core_ratio >= thresholds["min_core_ratio"]
            and nucleus_local_signal >= thresholds["min_nucleus_local_signal"]
            and nucleus_peak_signal >= thresholds["min_nucleus_peak_signal"]
            and (core_signal >= thresholds["min_core_signal"] or thick_ratio >= 0.025)
        )
        thick_absolute = (
            thick_ratio >= 0.055
            and nucleus_ratio >= 0.035
            and core_signal >= 4.0
            and nucleus_peak_signal >= thresholds["min_nucleus_peak_signal"]
        )
        color_absolute = (
            color_ratio >= thresholds["min_color_ratio"]
            and nucleus_color_ratio >= 0.025
            and color_signal >= 5.0
        )

        # Relative evidence recovers very light pencil ticks.  It is still
        # constrained to the bubble nucleus, so lighting gradients and print
        # weight differences do not qualify on contrast alone.
        relative_mark = (
            evidence >= thresholds["min_evidence"]
            and evidence_margin >= thresholds["min_relative_evidence"]
            and core_ratio >= 0.045
            and nucleus_ratio >= 0.028
            and nucleus_peak_signal >= 18.0
            and (nucleus_signal >= 3.0 or thick_ratio >= thresholds["min_thick_ratio"])
        )

        score["strength"] = float(
            evidence / thresholds["min_evidence"]
            + min(1.5, nucleus_ratio / max(thresholds["min_nucleus_ratio"], 1e-6))
            + min(1.0, thick_ratio / 0.055)
        ) / 3.0
        score["gate"] = (
            "color"
            if color_absolute
            else "dark"
            if dark_absolute
            else "thick"
            if thick_absolute
            else "relative"
            if relative_mark
            else "none"
        )
        if dark_absolute or thick_absolute or color_absolute or relative_mark:
            candidates.append(score)

    max_evidence = max(evidence_values)
    max_core_ratio = max(float(score.get("core_ink_ratio", 0.0)) for score in scores)
    max_nucleus_ratio = max(
        float(score.get("nucleus_ink_ratio", 0.0)) for score in scores
    )
    max_thick_ratio = max(
        float(score.get("core_thick_ratio", 0.0)) for score in scores
    )

    if candidates:
        selected = [str(score["option"]) for score in candidates]
        weakest = min(float(score["strength"]) for score in candidates)
        rejected = [score for score in scores if score not in candidates]
        strongest_rejected = max(
            (float(score["strength"]) for score in rejected),
            default=0.0,
        )
        separation = max(0.0, weakest - strongest_rejected)
        confidence = _clamp(
            0.62 + 0.16 * min(1.0, max(0.0, weakest - 1.0))
            + 0.18 * min(1.0, separation)
        )
        reason = "multiple" if len(selected) > 1 else "marked"
    else:
        selected = []
        clearly_blank = (
            max_evidence < thresholds["min_evidence"]
            and max_core_ratio < 0.07
            and max_nucleus_ratio < 0.055
            and max_thick_ratio < 0.035
        )
        reason = "blank" if clearly_blank else "ambiguous"
        if clearly_blank:
            distance = min(
                (thresholds["min_evidence"] - max_evidence)
                / thresholds["min_evidence"],
                (0.07 - max_core_ratio) / 0.07,
                (0.055 - max_nucleus_ratio) / 0.055,
            )
            confidence = _clamp(0.72 + 0.23 * max(0.0, distance))
        else:
            confidence = 0.38

    needs_review = not selected or len(selected) > 1 or confidence < 0.72
    return selected, {
        "reason": reason,
        "confidence": confidence,
        "needs_review": needs_review,
        "score_range": max(evidence_values) - min(evidence_values),
        "mean_combined": _mean(evidence_values),
        "std_combined": float(np.std(evidence_values)),
        "baselines": baselines,
        "thresholds": thresholds,
        "scores": scores,
    }


def select_filled_options(
    cell_scores: Sequence[Mapping[str, Any]],
) -> tuple[list[str], dict[str, Any]]:
    """Select marked cells using absolute evidence and a blank-cell baseline."""

    if not cell_scores:
        return [], {
            "reason": "invalid",
            "confidence": 0.0,
            "needs_review": True,
            "thresholds": {},
        }

    # New detections carry geometry-aware interior features.  Keep the older
    # score selector below for callers/tests that provide the historical score
    # dictionaries directly.
    if all("interior_evidence" in score for score in cell_scores):
        return _select_interior_scores(cell_scores)

    # Work on copies so callers can safely reuse their input dictionaries.
    scores = [dict(score) for score in cell_scores]
    combined_vals = [float(score["combined"]) for score in scores]
    max_combined = max(combined_vals)
    min_combined = min(combined_vals)
    score_range = max_combined - min_combined
    mean_combined = _mean(combined_vals)
    std_combined = float(np.std(combined_vals))

    sorted_scores = sorted(scores, key=lambda score: float(score["combined"]))
    baseline_count = max(1, len(sorted_scores) // 2)
    blank_reference = sorted_scores[:baseline_count]

    feature_names = (
        "combined",
        "darkness",
        "mark_ratio",
        "center_darkness",
        "center_mark_ratio",
        "interior_fill",
    )
    baselines = {
        feature: _mean([float(score.get(feature, 0.0)) for score in blank_reference])
        for feature in feature_names
    }
    maxima = {
        feature: max(float(score.get(feature, 0.0)) for score in scores)
        for feature in feature_names
    }

    for score in scores:
        for feature in feature_names:
            score[f"{feature}_margin"] = float(score.get(feature, 0.0)) - baselines[feature]

    min_combined = 4.5
    min_darkness = 1.8
    # Fixed evidence floors are deliberately independent of the strongest
    # option.  Scaling a threshold from the maximum caused a dark second mark
    # to make an existing faint mark disappear (non-monotonic multi-select).
    min_combined_margin = 3.2
    min_darkness_margin = 0.9
    min_mark_ratio = 0.035
    min_mark_margin = 0.016
    min_center_darkness = 2.6
    min_center_darkness_margin = 1.25
    min_center_mark_ratio = 0.065
    min_center_mark_margin = 0.012
    min_interior_fill = 0.025
    min_interior_fill_margin = 0.012

    thresholds = {
        "min_combined": min_combined,
        "min_darkness": min_darkness,
        "min_combined_margin": min_combined_margin,
        "min_darkness_margin": min_darkness_margin,
        "min_mark_ratio": min_mark_ratio,
        "min_mark_margin": min_mark_margin,
        "min_center_darkness": min_center_darkness,
        "min_center_darkness_margin": min_center_darkness_margin,
        "min_center_mark_ratio": min_center_mark_ratio,
        "min_center_mark_margin": min_center_mark_margin,
        "min_interior_fill": min_interior_fill,
        "min_interior_fill_margin": min_interior_fill_margin,
    }

    candidates: list[dict[str, Any]] = []
    for score in scores:
        passes_absolute = (
            float(score["combined"]) >= min_combined
            and float(score["darkness"]) >= min_darkness
        )
        passes_relative = (
            float(score["combined_margin"]) >= min_combined_margin
            and (
                float(score["darkness_margin"]) >= min_darkness_margin
                or float(score["center_darkness_margin"]) >= min_center_darkness_margin
            )
        )
        passes_ink = (
            float(score["mark_ratio"]) >= min_mark_ratio
            and float(score["mark_ratio_margin"]) >= min_mark_margin
        )
        passes_center = (
            (
                float(score["center_darkness"]) >= min_center_darkness
                and float(score["center_darkness_margin"]) >= min_center_darkness_margin
            )
            or (
                float(score["center_mark_ratio"]) >= min_center_mark_ratio
                and float(score["center_mark_ratio_margin"]) >= min_center_mark_margin
            )
            or (
                float(score["interior_fill"]) >= min_interior_fill
                and float(score["interior_fill_margin"]) >= min_interior_fill_margin
            )
        )
        # When every option is filled there is no blank cell from which to
        # derive a relative baseline.  Strong interior evidence is therefore
        # an independent path; printed outlines tend to live in the outer ring.
        strong_independent = (
            (
                float(score.get("interior_fill", 0.0)) >= 0.10
                and float(score.get("center_mark_ratio", 0.0)) >= 0.16
            )
            or (
                float(score.get("center_absolute_darkness", 0.0)) >= 18.0
                and float(score.get("center_mark_ratio", 0.0)) >= 0.12
            )
        )

        strength_parts = (
            float(score["combined_margin"]) / max(min_combined_margin, 1e-6),
            float(score["darkness_margin"]) / max(min_darkness_margin, 1e-6),
            float(score["mark_ratio_margin"]) / max(min_mark_margin, 1e-6),
            float(score["center_darkness_margin"]) / max(min_center_darkness_margin, 1e-6),
            float(score["center_mark_ratio_margin"]) / max(min_center_mark_margin, 1e-6),
            float(score["interior_fill_margin"]) / max(min_interior_fill_margin, 1e-6),
        )
        score["strength"] = float(
            0.30 * strength_parts[0]
            + 0.12 * strength_parts[1]
            + 0.16 * strength_parts[2]
            + 0.18 * strength_parts[3]
            + 0.14 * strength_parts[4]
            + 0.10 * strength_parts[5]
        )
        if (
            passes_absolute
            and passes_relative
            and (passes_ink or passes_center)
        ) or strong_independent:
            candidates.append(score)

    max_margin = max(float(score["combined_margin"]) for score in scores)
    clearly_blank = (
        score_range < 5.5
        and max_combined < 5.0
        and max_margin < 2.5
        and maxima["interior_fill"] < 0.055
        and maxima["center_mark_ratio"] < 0.12
    )

    if clearly_blank:
        blank_distance = min(
            (5.5 - score_range) / 5.5,
            (5.0 - max_combined) / 5.0,
            (2.5 - max_margin) / 2.5,
        )
        confidence = _clamp(0.68 + 0.28 * max(0.0, blank_distance))
        reason = "blank"
        selected: list[str] = []
    elif candidates:
        selected = [str(score["option"]) for score in candidates]
        selected_strength = min(float(score["strength"]) for score in candidates)
        unselected = [score for score in scores if score not in candidates]
        strongest_unselected = max(
            (float(score["strength"]) for score in unselected),
            default=0.0,
        )
        separation = selected_strength - strongest_unselected
        confidence = _clamp(
            0.52
            + 0.18 * min(1.5, max(0.0, selected_strength - 1.0))
            + 0.22 * min(1.0, max(0.0, separation)),
        )
        reason = "multiple" if len(selected) > 1 else "marked"
    else:
        selected = []
        top_strength = max(float(score["strength"]) for score in scores)
        # Some evidence exists, but it does not consistently clear both the
        # absolute and relative gates.  This is intentionally an uncertain
        # result instead of guessing the highest-scoring option.
        confidence = _clamp(0.50 - 0.18 * min(1.0, max(0.0, top_strength)))
        reason = "ambiguous" if max_combined >= min_combined else "blank"

    needs_review = not selected or len(selected) > 1 or confidence < 0.72
    decision = {
        "reason": reason,
        "confidence": confidence,
        "needs_review": needs_review,
        "score_range": score_range,
        "mean_combined": mean_combined,
        "std_combined": std_combined,
        "baselines": baselines,
        "thresholds": thresholds,
        "scores": scores,
    }
    return selected, decision


def detect_filled_options(image: Any, options_count: int = 4) -> OMRDetectionResult:
    """Detect filled choices in a cropped answer region.

    The crop should contain one horizontal row of equally spaced options.
    Blank and multiple responses are returned without guessing, and
    ``needs_review`` highlights answers that deserve a human check.
    """

    if not 2 <= int(options_count) <= len(OPTION_LABELS):
        raise ValueError(f"options_count must be between 2 and {len(OPTION_LABELS)}")

    rgb = _as_rgb_uint8(image)
    height, width = rgb.shape[:2]
    if height < 5 or width < options_count * 5:
        return OMRDetectionResult(
            answer="",
            confidence=0.0,
            needs_review=True,
            reason="invalid",
            cell_scores=(),
            content_bounds=(0, width),
            cell_edges=(0, width),
            decision={"reason": "invalid", "confidence": 0.0, "needs_review": True},
        )

    gray_u8 = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    gray = gray_u8.astype(np.float32)
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    saturation_u8 = hsv[:, :, 1]
    saturation = saturation_u8.astype(np.float32)

    # A high percentile estimates the paper colour more reliably than the
    # arithmetic mean, which is pulled down by a heavy mark.
    paper_level = float(np.percentile(gray, 78))
    low_level = float(np.percentile(gray, 5))
    contrast_range = max(1.0, paper_level - low_level)
    dark_pixel_threshold = paper_level - max(10.0, min(38.0, contrast_range * 0.20))
    saturation_median = float(np.median(saturation))
    saturation_std = float(np.std(saturation))
    color_pixel_threshold = saturation_median + max(18.0, saturation_std * 0.70)

    # A 3x3 median removes isolated sensor/salt-and-pepper noise before stroke
    # thickness is measured.  The morphology kernel is only used to estimate
    # the low-frequency colour floor once for the complete row; grayscale
    # paper reconstruction is performed locally per cell below.
    estimated_cell_width = max(5, int(round(width / options_count)))

    def odd_size(target: float, limit: int) -> int:
        value = max(3, min(int(round(target)), max(3, int(limit))))
        if value % 2 == 0:
            value -= 1
        return max(3, value)

    kernel_width = odd_size(estimated_cell_width * 0.58, width)
    kernel_height = odd_size(height * 0.68, height)
    background_kernel = cv2.getStructuringElement(
        cv2.MORPH_ELLIPSE,
        (kernel_width, kernel_height),
    )
    if min(height, width) >= 3:
        clean_gray = cv2.medianBlur(gray_u8, 3)
        clean_saturation = cv2.medianBlur(saturation_u8, 3)
    else:  # guarded by the size check above, kept for defensive reuse
        clean_gray = gray_u8
        clean_saturation = saturation_u8
    local_color_floor = cv2.morphologyEx(
        clean_saturation,
        cv2.MORPH_OPEN,
        background_kernel,
    ).astype(np.float32)
    color_response = np.maximum(
        clean_saturation.astype(np.float32) - local_color_floor,
        0.0,
    )

    # Template rectangles define the option group, so cell boundaries must be
    # fixed.  Inferring them again from the student's ink made the selected
    # option itself move the boundaries and could turn an empty row into a
    # false multi-select.
    content_left, content_right = 0, width
    content_span = max(1, content_right - content_left)
    cell_edges = [
        int(round(content_left + content_span * index / options_count))
        for index in range(options_count + 1)
    ]
    cell_edges[0] = content_left
    cell_edges[-1] = content_right
    rectangle_geometries = _estimate_repeated_rectangle_geometry(
        gray_u8,
        cell_edges,
    )

    enhanced = cv2.normalize(gray, None, 0, 255, cv2.NORM_MINMAX).astype(np.float32)
    overall_mean = float(np.mean(_safe_inner(gray, 0.02, 0.10)))
    overall_enhanced_mean = float(np.mean(_safe_inner(enhanced, 0.02, 0.10)))
    global_color_ratio = float(np.mean(saturation > color_pixel_threshold))

    cell_scores: list[dict[str, Any]] = []
    for index in range(options_count):
        left, right = cell_edges[index], cell_edges[index + 1]
        cell_gray = gray[:, left:right]
        cell_enhanced = enhanced[:, left:right]
        cell_saturation = saturation[:, left:right]
        cell_rgb = rgb[:, left:right]
        cell_color_response = color_response[:, left:right]

        # Excluding a narrow outer ring prevents table/grid borders from being
        # mistaken for handwriting while retaining ticks and partial fills.
        body_gray = _safe_inner(cell_gray, 0.07, 0.12)
        body_enhanced = _safe_inner(cell_enhanced, 0.07, 0.12)
        body_saturation = _safe_inner(cell_saturation, 0.07, 0.12)
        center_gray = _safe_inner(cell_gray, 0.24, 0.22)
        center_saturation = _safe_inner(cell_saturation, 0.24, 0.22)

        body_dark_mask = body_gray < dark_pixel_threshold
        center_dark_mask = center_gray < dark_pixel_threshold
        body_color_mask = body_saturation > color_pixel_threshold
        center_color_mask = center_saturation > color_pixel_threshold

        cell_height, cell_width = cell_gray.shape[:2]
        yy, xx = np.ogrid[:cell_height, :cell_width]
        if rectangle_geometries is not None:
            center_x, center_y, target_width, target_height = (
                rectangle_geometries[index]
            )
            support_rx = max(2.0, target_width * 0.50)
            support_ry = max(2.0, target_height * 0.50)
            target_source = "repeated_rectangle"

            # Replace the bubble ellipses with the measured writable interior.
            # Float boundaries follow OpenCV's inclusive bounding rectangle
            # convention and keep a strict two-pixel distance from print.
            target_left = center_x - (target_width - 1.0) * 0.50
            target_right = center_x + (target_width - 1.0) * 0.50
            target_top = center_y - (target_height - 1.0) * 0.50
            target_bottom = center_y + (target_height - 1.0) * 0.50
            target_box_mask = (
                (xx >= target_left - 0.5)
                & (xx <= target_right + 0.5)
                & (yy >= target_top - 0.5)
                & (yy <= target_bottom + 0.5)
            )
            core_mask = (
                (xx >= target_left + 2.0)
                & (xx <= target_right - 2.0)
                & (yy >= target_top + 2.0)
                & (yy <= target_bottom - 2.0)
            )
            if not np.any(core_mask):
                core_mask = (
                    (xx >= target_left + 1.0)
                    & (xx <= target_right - 1.0)
                    & (yy >= target_top + 1.0)
                    & (yy <= target_bottom - 1.0)
                )
            nucleus_mask = (
                core_mask
                & (xx >= center_x - target_width * 0.27)
                & (xx <= center_x + target_width * 0.27)
            )
            outline_ring_mask = target_box_mask & np.logical_not(core_mask)
            outer_paper_mask = (
                (xx >= target_left - 4.0)
                & (xx <= target_right + 4.0)
                & (yy >= target_top - 4.0)
                & (yy <= target_bottom + 4.0)
                & np.logical_not(target_box_mask)
            )
            if not np.any(outer_paper_mask):
                outer_paper_mask = np.logical_not(core_mask)
        else:
            center_x = (cell_width - 1.0) / 2.0
            center_y = (cell_height - 1.0) / 2.0
            support_rx = max(2.0, min(cell_width * 0.27, cell_height * 0.36))
            support_ry = max(2.0, min(cell_height * 0.36, cell_width * 0.27))
            target_width = support_rx * 2.0
            target_height = support_ry * 2.0
            target_source = "crop_center"
            radius_squared = (
                ((xx - center_x) / support_rx) ** 2
                + ((yy - center_y) / support_ry) ** 2
            )
            core_mask = radius_squared <= 0.58**2
            nucleus_mask = radius_squared <= 0.31**2
            outline_ring_mask = np.logical_and(
                radius_squared >= 0.70**2,
                radius_squared <= 1.12**2,
            )
            outer_paper_mask = np.logical_and(
                radius_squared >= 1.34**2,
                radius_squared <= 1.85**2,
            )
            if not np.any(outer_paper_mask):
                outer_paper_mask = radius_squared >= 1.24**2

        # Estimate paper from a symmetric ring outside the printed bubble.
        # This is both much faster than per-cell inpainting and insensitive to
        # a linear scanner gradient: opposite sides balance at the centre.
        # It remains valid when every option is filled because the surrounding
        # paper in each individual cell is still unmarked.
        if target_source == "repeated_rectangle":
            # Preserve one-pixel pencil/blue strokes.  The strict rectangle
            # erosion supplies the print/noise guard that median blur provides
            # for crop-centred bubbles.
            cell_clean_gray = gray_u8[:, left:right]
            outer_values = cell_clean_gray[outer_paper_mask]
            local_paper_level = float(np.percentile(outer_values, 70))
            saturation_floor = float(np.median(cell_saturation[outer_paper_mask]))
            cell_color_response = np.maximum(
                cell_saturation - saturation_floor,
                0.0,
            )
            thick_distance_threshold = 1.35
        else:
            cell_clean_gray = clean_gray[:, left:right]
            outer_values = cell_clean_gray[outer_paper_mask]
            local_paper_level = float(np.percentile(outer_values, 65))
            thick_distance_threshold = 1.85
        cell_dark_response = np.maximum(
            local_paper_level - cell_clean_gray.astype(np.float32),
            0.0,
        )

        robust_ink_mask = np.logical_or(
            cell_dark_response >= 10.0,
            np.logical_and(cell_color_response >= 20.0, cell_saturation >= 32.0),
        )
        distance = cv2.distanceTransform(
            robust_ink_mask.astype(np.uint8),
            cv2.DIST_L2,
            3,
        )
        thick_ink_mask = distance >= thick_distance_threshold

        def masked_mean(values: np.ndarray, mask: np.ndarray) -> float:
            return float(np.mean(values[mask])) if np.any(mask) else 0.0

        def masked_ratio(mask: np.ndarray, region: np.ndarray) -> float:
            return float(np.mean(mask[region])) if np.any(region) else 0.0

        core_dark_signal = masked_mean(cell_dark_response, core_mask)
        nucleus_dark_signal = masked_mean(cell_dark_response, nucleus_mask)
        ring_dark_signal = masked_mean(cell_dark_response, outline_ring_mask)
        core_ink_ratio = masked_ratio(robust_ink_mask, core_mask)
        nucleus_ink_ratio = masked_ratio(robust_ink_mask, nucleus_mask)
        core_thick_ratio = masked_ratio(thick_ink_mask, core_mask)
        core_color_signal = masked_mean(cell_color_response, core_mask)
        nucleus_color_signal = masked_mean(cell_color_response, nucleus_mask)
        core_color_ratio = masked_ratio(
            np.logical_and(cell_color_response >= 20.0, cell_saturation >= 32.0),
            core_mask,
        )
        nucleus_color_ratio = masked_ratio(
            np.logical_and(cell_color_response >= 20.0, cell_saturation >= 32.0),
            nucleus_mask,
        )

        nucleus_values = cell_clean_gray[nucleus_mask].astype(np.float32)
        core_values = cell_clean_gray[core_mask].astype(np.float32)
        nucleus_local_signal = max(
            0.0,
            local_paper_level - float(np.mean(nucleus_values)),
        )
        # A lower percentile retains thin ticks without reacting to one or two
        # isolated noise pixels (which the median filter has already removed).
        nucleus_peak_signal = max(
            0.0,
            local_paper_level - float(np.percentile(nucleus_values, 35)),
        )
        core_local_signal = max(
            0.0,
            local_paper_level - float(np.mean(core_values)),
        )

        interior_evidence = (
            core_dark_signal * 0.11
            + nucleus_dark_signal * 0.08
            + core_ink_ratio * 12.0
            + nucleus_ink_ratio * 10.0
            + core_thick_ratio * 15.0
            + core_color_signal * 0.06
            + nucleus_color_signal * 0.04
            + core_color_ratio * 12.0
            + nucleus_color_ratio * 8.0
            + nucleus_local_signal * 0.08
            + nucleus_peak_signal * 0.05
        )

        mark_ratio = max(
            float(np.mean(body_dark_mask)),
            float(np.mean(body_color_mask)),
        )
        center_mark_ratio = max(
            float(np.mean(center_dark_mask)),
            float(np.mean(center_color_mask)),
        )
        interior_fill = max(core_ink_ratio, core_color_ratio)

        cell_mean = float(np.mean(body_gray))
        center_mean = float(np.mean(center_gray))
        darkness = overall_mean - cell_mean
        center_darkness = overall_mean - center_mean
        absolute_darkness = paper_level - cell_mean
        center_absolute_darkness = paper_level - center_mean
        enhanced_darkness = overall_enhanced_mean - float(np.mean(body_enhanced))
        local_contrast = float(np.std(body_gray)) / 10.0

        color_ratio = float(np.mean(body_color_mask))
        center_color_ratio = float(np.mean(center_color_mask))
        saturation_score = float(np.mean(body_saturation)) - float(np.mean(saturation))
        # Blue ink is common on school answer sheets.  HSV saturation handles
        # every hue; this small channel term helps blue marks that are light.
        blue_score = float(
            np.mean(
                cell_rgb[:, :, 2].astype(np.float32)
                - cell_rgb[:, :, 0].astype(np.float32)
            )
        )
        color_excess = max(0.0, color_ratio - global_color_ratio)

        # ``combined`` remains for compatibility with existing diagnostics;
        # decisions use the explicitly named interior evidence above.
        combined = interior_evidence
        cell_scores.append(
            {
                "option": OPTION_LABELS[index],
                "gray_mean": cell_mean,
                "darkness": darkness,
                "center_darkness": center_darkness,
                "absolute_darkness": absolute_darkness,
                "center_absolute_darkness": center_absolute_darkness,
                "enhanced_dark": enhanced_darkness,
                "local_contrast": local_contrast,
                "mark_ratio": mark_ratio,
                "center_mark_ratio": center_mark_ratio,
                "interior_fill": interior_fill,
                "color_ratio": color_ratio,
                "center_color_ratio": center_color_ratio,
                "blue_score": blue_score,
                "core_dark_signal": core_dark_signal,
                "nucleus_dark_signal": nucleus_dark_signal,
                "ring_dark_signal": ring_dark_signal,
                "core_ink_ratio": core_ink_ratio,
                "nucleus_ink_ratio": nucleus_ink_ratio,
                "core_thick_ratio": core_thick_ratio,
                "core_color_signal": core_color_signal,
                "nucleus_color_signal": nucleus_color_signal,
                "core_color_ratio": core_color_ratio,
                "nucleus_color_ratio": nucleus_color_ratio,
                "local_paper_level": local_paper_level,
                "target_source": target_source,
                "target_center_x": center_x,
                "target_center_y": center_y,
                "target_radius_x": support_rx,
                "target_radius_y": support_ry,
                "core_local_signal": core_local_signal,
                "nucleus_local_signal": nucleus_local_signal,
                "nucleus_peak_signal": nucleus_peak_signal,
                "interior_evidence": interior_evidence,
                "combined": combined,
            }
        )

    selected, decision = select_filled_options(cell_scores)
    # Pull the enriched score copies out of the decision structure.
    scored_cells = tuple(decision.pop("scores", cell_scores))
    answer = "".join(dict.fromkeys(selected))
    return OMRDetectionResult(
        answer=answer,
        confidence=float(decision["confidence"]),
        needs_review=bool(decision["needs_review"]),
        reason=str(decision["reason"]),
        cell_scores=scored_cells,
        content_bounds=(content_left, content_right),
        cell_edges=tuple(cell_edges),
        decision=decision,
    )


def estimate_skew_angle(image: Any, max_abs_angle: float = 6.0) -> float:
    """Estimate a small page skew using length-weighted horizontal lines."""

    rgb = _as_rgb_uint8(image)
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    height, width = gray.shape[:2]

    scale = min(1.0, 1600.0 / max(height, width))
    if scale < 1.0:
        gray = cv2.resize(
            gray,
            (max(1, int(round(width * scale))), max(1, int(round(height * scale)))),
            interpolation=cv2.INTER_AREA,
        )

    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    median = float(np.median(blurred))
    lower = int(max(20, 0.66 * median))
    upper = int(min(250, max(lower + 20, 1.33 * median)))
    edges = cv2.Canny(blurred, lower, upper, apertureSize=3)
    min_dimension = min(gray.shape[:2])
    lines = cv2.HoughLinesP(
        edges,
        1,
        np.pi / 1800.0,
        threshold=max(45, int(min_dimension * 0.06)),
        minLineLength=max(55, int(gray.shape[1] * 0.12)),
        maxLineGap=max(8, int(gray.shape[1] * 0.015)),
    )
    if lines is None:
        return 0.0

    observations: list[tuple[float, float]] = []
    for x1, y1, x2, y2 in lines[:, 0]:
        dx = float(x2 - x1)
        dy = float(y2 - y1)
        length = float(np.hypot(dx, dy))
        if length <= 0.0:
            continue
        angle = float(np.degrees(np.arctan2(dy, dx)))
        if abs(angle) <= max_abs_angle:
            observations.append((angle, length))
    if not observations:
        return 0.0

    observations.sort(key=lambda item: item[0])
    total_weight = sum(weight for _, weight in observations)
    midpoint = total_weight / 2.0
    cumulative = 0.0
    weighted_median = 0.0
    for angle, weight in observations:
        cumulative += weight
        if cumulative >= midpoint:
            weighted_median = angle
            break

    deviations = np.array([abs(angle - weighted_median) for angle, _ in observations])
    mad = float(np.median(deviations))
    tolerance = max(0.35, mad * 3.5)
    inliers = [
        (angle, weight)
        for angle, weight in observations
        if abs(angle - weighted_median) <= tolerance
    ]
    if not inliers:
        return 0.0
    angle = sum(value * weight for value, weight in inliers) / sum(
        weight for _, weight in inliers
    )
    return float(angle) if 0.15 <= abs(angle) <= max_abs_angle else 0.0


def deskew_image(image: Any) -> tuple[np.ndarray, float]:
    """Deskew a page while preserving its original pixel coordinate system."""

    rgb = _as_rgb_uint8(image)
    angle = estimate_skew_angle(rgb)
    if angle == 0.0:
        return rgb, 0.0

    height, width = rgb.shape[:2]
    matrix = cv2.getRotationMatrix2D((width / 2.0, height / 2.0), angle, 1.0)
    corrected = cv2.warpAffine(
        rgb,
        matrix,
        (width, height),
        flags=cv2.INTER_LINEAR,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(255, 255, 255),
    )
    return corrected, angle


def estimate_similarity_transform(
    source_points: Sequence[Sequence[float]],
    target_points: Sequence[Sequence[float]],
    *,
    ransac_threshold: float = 3.0,
) -> tuple[np.ndarray | None, dict[str, Any]]:
    """Estimate a robust current-page → reference-page similarity transform."""

    source = np.asarray(source_points, dtype=np.float32).reshape(-1, 2)
    target = np.asarray(target_points, dtype=np.float32).reshape(-1, 2)
    if len(source) != len(target) or len(source) < 2:
        return None, {
            "reason": "not_enough_points",
            "inliers": 0,
            "rms_error": float("inf"),
            "scale": 1.0,
            "rotation": 0.0,
        }

    matrix, inlier_mask = cv2.estimateAffinePartial2D(
        source,
        target,
        method=cv2.RANSAC,
        ransacReprojThreshold=float(ransac_threshold),
        maxIters=3000,
        confidence=0.995,
        refineIters=20,
    )
    if matrix is None:
        return None, {
            "reason": "estimation_failed",
            "inliers": 0,
            "rms_error": float("inf"),
            "scale": 1.0,
            "rotation": 0.0,
        }

    inliers = (
        np.ones(len(source), dtype=bool)
        if inlier_mask is None
        else inlier_mask.reshape(-1).astype(bool)
    )
    inlier_count = int(np.count_nonzero(inliers))
    projected = cv2.transform(source.reshape(1, -1, 2), matrix).reshape(-1, 2)
    residuals = np.linalg.norm(projected - target, axis=1)
    rms_error = float(
        np.sqrt(np.mean(np.square(residuals[inliers])))
        if inlier_count
        else float("inf")
    )
    a, b = float(matrix[0, 0]), float(matrix[0, 1])
    scale = float(np.hypot(a, b))
    rotation = float(np.degrees(np.arctan2(-b, a)))
    valid = (
        inlier_count >= 2
        and 0.90 <= scale <= 1.10
        and rms_error <= max(2.5, float(ransac_threshold))
    )
    details = {
        "reason": "ok" if valid else "quality_rejected",
        "inliers": inlier_count,
        "total_points": len(source),
        "rms_error": rms_error,
        "scale": scale,
        "rotation": rotation,
    }
    return (matrix.astype(np.float32), details) if valid else (None, details)
