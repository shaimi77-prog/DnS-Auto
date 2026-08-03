"""페이지 전처리 코어의 게이트·상태 격리 회귀시험."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

import unittest

import cv2
import numpy as np

from page_preprocessing import (
    OrientationPrediction,
    PagePreprocessConfig,
    PagePreprocessor,
    estimate_skew,
    inverse_transform_box,
    orientation_prediction,
    rotate_fine,
)


def lined_page(angle=0.0):
    image = np.full((700, 900, 3), 255, dtype=np.uint8)
    for index in range(12):
        y = 90 + index * 42
        cv2.line(image, (100, y), (790, y), (0, 0, 0), 5)
        for x in range(120, 760, 90):
            cv2.rectangle(image, (x, y - 18), (x + 45, y - 4), (0, 0, 0), -1)
    if not angle:
        return image
    rotated, _matrix = rotate_fine(image, angle)
    return rotated


class PagePreprocessingTests(unittest.TestCase):
    def test_disabled_returns_copy_without_transform(self):
        image = lined_page()
        result = PagePreprocessor().preprocess(image, pdf_rotation=90)
        self.assertEqual(result.status, "unchanged")
        self.assertEqual(result.pdf_rotation, 90)
        self.assertFalse(result.orientation_applied)
        self.assertFalse(result.deskew_applied)
        self.assertIsNot(result.processed_image, image)
        np.testing.assert_array_equal(result.transform_matrix, np.eye(3))

    def test_orientation_gate_requires_confidence_and_margin(self):
        config = PagePreprocessConfig(enabled=True, orientation_enabled=True)
        strong = PagePreprocessor(
            config,
            lambda _image: OrientationPrediction(90, 0.97, 0.22),
        ).preprocess(lined_page())
        weak = PagePreprocessor(
            config,
            lambda _image: OrientationPrediction(90, 0.91, 0.30),
        ).preprocess(lined_page())
        close = PagePreprocessor(
            config,
            lambda _image: OrientationPrediction(90, 0.97, 0.10),
        ).preprocess(lined_page())
        self.assertTrue(strong.orientation_applied)
        self.assertEqual(strong.processed_image.shape[:2], (900, 700))
        self.assertFalse(weak.orientation_applied)
        self.assertFalse(close.orientation_applied)

    def test_probability_helper_reports_margin(self):
        prediction = orientation_prediction({0: 0.05, 90: 0.91, 180: 0.03, 270: 0.01})
        self.assertEqual(prediction.degrees, 90)
        self.assertAlmostEqual(prediction.confidence, 0.91)
        self.assertAlmostEqual(prediction.margin, 0.86)

    def test_skew_estimation_and_correction(self):
        image = lined_page(3.0)
        before = estimate_skew(image)
        self.assertGreaterEqual(before.valid_line_count, 5)
        self.assertAlmostEqual(abs(before.angle), 3.0, delta=0.8)
        config = PagePreprocessConfig(enabled=True, deskew_enabled=True)
        result = PagePreprocessor(config).preprocess(image)
        self.assertTrue(result.deskew_applied)
        after = estimate_skew(result.processed_image)
        self.assertLess(abs(after.angle), 0.7)

    def test_blank_page_is_not_corrected(self):
        blank = np.full((500, 700, 3), 255, dtype=np.uint8)
        config = PagePreprocessConfig(enabled=True, deskew_enabled=True)
        result = PagePreprocessor(config).preprocess(blank)
        self.assertEqual(result.status, "unchanged")
        self.assertEqual(result.valid_line_count, 0)
        self.assertFalse(result.deskew_applied)

    def test_page_results_do_not_leak_between_calls(self):
        calls = iter(
            [
                OrientationPrediction(90, 0.99, 0.30),
                OrientationPrediction(0, 0.99, 0.30),
            ]
        )
        preprocessor = PagePreprocessor(
            PagePreprocessConfig(enabled=True, orientation_enabled=True),
            lambda _image: next(calls),
        )
        first = preprocessor.preprocess(lined_page())
        second = preprocessor.preprocess(lined_page())
        self.assertTrue(first.orientation_applied)
        self.assertFalse(second.orientation_applied)
        self.assertEqual(second.detected_orientation, 0)
        np.testing.assert_array_equal(second.transform_matrix, np.eye(3))
        self.assertIsNot(first.processed_image, second.processed_image)

    def test_reference_rejection_is_page_local(self):
        config = PagePreprocessConfig(enabled=True, orientation_enabled=True)
        preprocessor = PagePreprocessor(
            config,
            lambda _image: OrientationPrediction(90, 0.99, 0.30),
        )
        rejected = preprocessor.preprocess(
            lined_page(), reference_validator=lambda _result: (False, "anchor_mismatch")
        )
        accepted = preprocessor.preprocess(
            lined_page(), reference_validator=lambda _result: True
        )
        self.assertEqual(rejected.status, "rejected")
        self.assertEqual(rejected.failure_reason, "anchor_mismatch")
        self.assertEqual(accepted.reference_validation, "accepted")
        self.assertEqual(accepted.status, "corrected")

    def test_inverse_transform_restores_original_box(self):
        image = lined_page()
        rotated, transform = rotate_fine(image, 4.0)
        self.assertGreater(rotated.shape[0], image.shape[0])
        original = np.array(
            [[120.0, 80.0], [320.0, 80.0], [320.0, 140.0], [120.0, 140.0]]
        )
        homogeneous = np.column_stack([original, np.ones(4)])
        transformed = (transform @ homogeneous.T).T[:, :2]
        restored = inverse_transform_box(transformed, transform)
        np.testing.assert_allclose(restored, original, atol=1e-6)

if __name__ == "__main__":
    unittest.main()
