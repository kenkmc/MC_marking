from __future__ import annotations

import unittest

import cv2
import numpy as np

from omr_core import (
    deskew_image,
    detect_filled_options,
    estimate_similarity_transform,
    estimate_skew_angle,
    select_filled_options,
)


def make_answer_region(
    marked=(),
    *,
    fill_value=35,
    fill_color=None,
    width=160,
    height=32,
    options_count=4,
):
    image = np.full((height, width, 3), 248, dtype=np.uint8)
    cell_width = width // options_count
    for index in range(options_count):
        center_x = index * cell_width + cell_width // 2
        cv2.circle(image, (center_x, height // 2), 8, (115, 115, 115), 1)
        if index in marked:
            color = fill_color or (fill_value, fill_value, fill_value)
            cv2.circle(image, (center_x, height // 2), 6, color, -1)
    return image


def make_degraded_answer_region(
    marked=(),
    *,
    seed=0,
    fill_value=155,
    fill_color=None,
    noise_sigma=7.0,
    salt_ratio=0.006,
):
    """Create a deterministic scan-like row with gradients and uneven print."""

    rng = np.random.default_rng(seed)
    height, width = 36, 176
    x_gradient = np.linspace(-22.0, 18.0, width, dtype=np.float32)[None, :]
    y_gradient = np.linspace(7.0, -5.0, height, dtype=np.float32)[:, None]
    paper = 232.0 + x_gradient + y_gradient

    # A smooth local scanner shadow is deliberately centred on one blank
    # option.  It should not behave like ink after local background removal.
    yy, xx = np.mgrid[:height, :width]
    shadow = 18.0 * np.exp(-(((xx - 66.0) / 27.0) ** 2 + ((yy - 18.0) / 18.0) ** 2))
    paper -= shadow
    image = np.repeat(paper[:, :, None], 3, axis=2)
    image = np.clip(image, 0, 255).astype(np.uint8)

    cell_width = width // 4
    outline_thickness = (1, 3, 2, 1)
    radii = (8, 9, 8, 9)
    offsets = ((-1, 0), (1, -1), (0, 1), (1, 0))
    for index in range(4):
        center = (
            index * cell_width + cell_width // 2 + offsets[index][0],
            height // 2 + offsets[index][1],
        )
        outline_gray = 92 + index * 8
        cv2.circle(
            image,
            center,
            radii[index],
            (outline_gray, outline_gray, outline_gray),
            outline_thickness[index],
        )
        if index in marked:
            color = fill_color or (fill_value, fill_value, fill_value)
            cv2.circle(image, center, 6, color, -1)

    gaussian = rng.normal(0.0, noise_sigma, image.shape[:2])[:, :, None]
    image = np.clip(image.astype(np.float32) + gaussian, 0, 255).astype(np.uint8)
    salt_count = int(round(height * width * salt_ratio))
    if salt_count:
        positions = rng.choice(height * width, salt_count, replace=False)
        rows, columns = np.unravel_index(positions, (height, width))
        values = rng.choice(np.array([0, 255], dtype=np.uint8), salt_count)
        image[rows, columns] = values[:, None]
    return image


def score(option, combined, darkness, mark, *, center=None, interior=0.03):
    center_value = darkness if center is None else center
    return {
        "option": option,
        "combined": combined,
        "darkness": darkness,
        "mark_ratio": mark,
        "center_darkness": center_value,
        "center_mark_ratio": mark,
        "interior_fill": interior,
        "center_absolute_darkness": max(0.0, center_value + 5.0),
    }


class OMRDetectionTests(unittest.TestCase):
    def test_blank_region_is_not_guessed(self):
        result = detect_filled_options(make_answer_region())
        self.assertEqual(result.answer, "")
        self.assertTrue(result.needs_review)

    def test_single_dark_mark(self):
        result = detect_filled_options(make_answer_region((1,)))
        self.assertEqual(result.answer, "B")
        self.assertGreaterEqual(result.confidence, 0.60)

    def test_faint_pencil_mark(self):
        result = detect_filled_options(make_answer_region((2,), fill_value=150))
        self.assertEqual(result.answer, "C")

    def test_faint_check_mark_is_not_limited_to_solid_fills(self):
        image = make_answer_region()
        cv2.line(image, (55, 16), (59, 20), (120, 120, 120), 2)
        cv2.line(image, (59, 20), (68, 10), (120, 120, 120), 2)
        result = detect_filled_options(image)
        self.assertEqual(result.answer, "B")

    def test_blue_pen_mark(self):
        # Input arrays use RGB ordering.
        result = detect_filled_options(
            make_answer_region((0,), fill_color=(35, 80, 220))
        )
        self.assertEqual(result.answer, "A")

    def test_multiple_marks_are_preserved(self):
        result = detect_filled_options(make_answer_region((1, 3)))
        self.assertEqual(result.answer, "BD")
        self.assertTrue(result.needs_review)

    def test_non_default_option_count_uses_the_same_geometry(self):
        image = make_answer_region(
            (1, 4), width=200, height=34, options_count=5
        )
        result = detect_filled_options(image, options_count=5)
        self.assertEqual(result.answer, "BE")

    def test_all_marked_does_not_become_blank(self):
        result = detect_filled_options(make_answer_region((0, 1, 2, 3)))
        self.assertEqual(result.answer, "ABCD")
        self.assertTrue(result.needs_review)

    def test_mark_position_does_not_move_cell_boundaries(self):
        first = detect_filled_options(make_answer_region((0,)))
        last = detect_filled_options(make_answer_region((3,)))
        self.assertEqual(first.cell_edges, last.cell_edges)

    def test_adding_a_strong_mark_does_not_hide_a_faint_mark(self):
        blanks = [
            score("A", 0.0, 0.0, 0.0, interior=0.0),
            score("C", 0.0, 0.0, 0.0, interior=0.0),
            score("D", 0.0, 0.0, 0.0, interior=0.0),
        ]
        faint_b = score("B", 6.0, 4.0, 0.10, center=4.0)
        selected, _ = select_filled_options([blanks[0], faint_b, *blanks[1:]])
        self.assertIn("B", selected)

        heavy_a = score("A", 20.0, 15.0, 0.30, center=15.0)
        selected_with_heavy, _ = select_filled_options(
            [heavy_a, faint_b, *blanks[1:]]
        )
        self.assertIn("A", selected_with_heavy)
        self.assertIn("B", selected_with_heavy)

    def test_noisy_gradient_blanks_with_uneven_outlines_stay_blank(self):
        # Multiple seeds cover different isolated noise positions without
        # making this a probabilistic/flaky test.
        for seed in range(20):
            with self.subTest(seed=seed):
                result = detect_filled_options(
                    make_degraded_answer_region(seed=seed)
                )
                self.assertEqual(result.answer, "")
                self.assertIn(result.reason, {"blank", "ambiguous"})

    def test_hollow_heavy_print_is_not_a_filled_response(self):
        image = np.full((36, 176, 3), 242, dtype=np.uint8)
        cell_width = image.shape[1] // 4
        for index in range(4):
            center_x = index * cell_width + cell_width // 2
            # Mix square and circular, unusually heavy printed outlines.
            if index % 2:
                cv2.rectangle(
                    image,
                    (center_x - 8, 10),
                    (center_x + 8, 26),
                    (70, 70, 70),
                    3,
                )
            else:
                cv2.circle(image, (center_x, 18), 9, (70, 70, 70), 3)

        result = detect_filled_options(image)
        self.assertEqual(result.answer, "")

    def test_faint_mark_survives_scan_gradient_and_noise(self):
        for seed in range(8):
            with self.subTest(seed=seed):
                result = detect_filled_options(
                    make_degraded_answer_region(
                        (2,),
                        seed=100 + seed,
                        fill_value=165,
                        noise_sigma=9.0,
                    )
                )
                self.assertEqual(result.answer, "C")

    def test_blue_and_multiple_marks_survive_scan_degradation(self):
        blue = detect_filled_options(
            make_degraded_answer_region(
                (0,),
                seed=301,
                fill_color=(60, 110, 205),
                noise_sigma=8.0,
            )
        )
        self.assertEqual(blue.answer, "A")

        multiple = detect_filled_options(
            make_degraded_answer_region((1, 3), seed=302, noise_sigma=8.0)
        )
        self.assertEqual(multiple.answer, "BD")

    def test_all_marks_survive_scan_degradation(self):
        result = detect_filled_options(
            make_degraded_answer_region((0, 1, 2, 3), seed=410, noise_sigma=8.0)
        )
        self.assertEqual(result.answer, "ABCD")

    def test_image_level_multi_select_is_monotonic(self):
        faint_only = make_degraded_answer_region(
            (1,), seed=777, fill_value=172, noise_sigma=6.0, salt_ratio=0.0
        )
        faint_and_heavy = faint_only.copy()
        cell_width = faint_and_heavy.shape[1] // 4
        cv2.circle(
            faint_and_heavy,
            (cell_width // 2 - 1, faint_and_heavy.shape[0] // 2),
            6,
            (25, 25, 25),
            -1,
        )

        first = detect_filled_options(faint_only)
        second = detect_filled_options(faint_and_heavy)
        self.assertIn("B", first.answer)
        self.assertIn("A", second.answer)
        self.assertIn("B", second.answer)

    def test_resolution_blur_and_jpeg_artifacts_do_not_change_answer(self):
        marked = make_degraded_answer_region(
            (3,), seed=901, fill_value=170, noise_sigma=5.0, salt_ratio=0.0
        )
        blank = make_degraded_answer_region(
            seed=902, noise_sigma=5.0, salt_ratio=0.0
        )
        for scale in (0.75, 1.0, 1.5, 2.0):
            with self.subTest(scale=scale):
                size = (
                    int(round(marked.shape[1] * scale)),
                    int(round(marked.shape[0] * scale)),
                )
                scaled_marked = cv2.resize(marked, size, interpolation=cv2.INTER_LINEAR)
                scaled_blank = cv2.resize(blank, size, interpolation=cv2.INTER_LINEAR)
                scaled_marked = cv2.GaussianBlur(scaled_marked, (3, 3), 0.55)
                scaled_blank = cv2.GaussianBlur(scaled_blank, (3, 3), 0.55)

                # Simulate a low-quality PDF/JPEG scan round trip.
                ok, encoded_marked = cv2.imencode(
                    ".jpg",
                    cv2.cvtColor(scaled_marked, cv2.COLOR_RGB2BGR),
                    [cv2.IMWRITE_JPEG_QUALITY, 48],
                )
                self.assertTrue(ok)
                ok, encoded_blank = cv2.imencode(
                    ".jpg",
                    cv2.cvtColor(scaled_blank, cv2.COLOR_RGB2BGR),
                    [cv2.IMWRITE_JPEG_QUALITY, 48],
                )
                self.assertTrue(ok)
                decoded_marked = cv2.cvtColor(
                    cv2.imdecode(encoded_marked, cv2.IMREAD_COLOR),
                    cv2.COLOR_BGR2RGB,
                )
                decoded_blank = cv2.cvtColor(
                    cv2.imdecode(encoded_blank, cv2.IMREAD_COLOR),
                    cv2.COLOR_BGR2RGB,
                )
                self.assertEqual(detect_filled_options(decoded_marked).answer, "D")
                self.assertEqual(detect_filled_options(decoded_blank).answer, "")


class DeskewTests(unittest.TestCase):
    def test_deskew_preserves_canvas_and_reduces_angle(self):
        original = np.full((900, 650, 3), 255, dtype=np.uint8)
        for y in range(120, 820, 90):
            cv2.line(original, (70, y), (580, y), (25, 25, 25), 3)
        matrix = cv2.getRotationMatrix2D((325, 450), 2.0, 1.0)
        skewed = cv2.warpAffine(
            original,
            matrix,
            (650, 900),
            borderValue=(255, 255, 255),
        )

        before = estimate_skew_angle(skewed)
        corrected, applied = deskew_image(skewed)
        after = estimate_skew_angle(corrected)

        self.assertEqual(corrected.shape, skewed.shape)
        self.assertGreater(abs(before), 1.4)
        self.assertGreater(abs(applied), 1.4)
        self.assertLess(abs(after), 0.45)

    def test_similarity_transform_rejects_one_bad_anchor(self):
        source = np.array(
            [[80.0, 90.0], [560.0, 85.0], [75.0, 790.0], [565.0, 795.0]],
            dtype=np.float32,
        )
        expected = cv2.getRotationMatrix2D((325.0, 450.0), 1.2, 1.015)
        expected[:, 2] += np.array([13.0, -9.0])
        target = cv2.transform(source.reshape(1, -1, 2), expected).reshape(-1, 2)
        target[3] += np.array([45.0, -30.0])

        matrix, quality = estimate_similarity_transform(source, target)
        self.assertIsNotNone(matrix)
        self.assertGreaterEqual(quality["inliers"], 3)
        projected = cv2.transform(
            source[:3].reshape(1, -1, 2), matrix
        ).reshape(-1, 2)
        self.assertLess(float(np.max(np.linalg.norm(projected - target[:3], axis=1))), 1.0)


if __name__ == "__main__":
    unittest.main()
