"""OCR 전처리 이미지 좌표가 원래 PDF 좌표로 복원되는지 검증."""

import unittest

import fitz
import numpy as np

from engine_Drag import HybridTextExtractor
from page_preprocessing import rotate_fine


class OcrPreprocessMappingTests(unittest.TestCase):
    def test_transformed_box_returns_to_original_pdf_coordinates(self):
        image = np.full((1000, 800, 3), 255, dtype=np.uint8)
        _rotated, transform = rotate_fine(image, 3.5)
        original_pixels = np.array(
            [[160.0, 200.0], [400.0, 200.0], [400.0, 300.0], [160.0, 300.0]]
        )
        homogeneous = np.column_stack([original_pixels, np.ones(4)])
        transformed_box = (transform @ homogeneous.T).T[:, :2]
        pdf_rect = HybridTextExtractor._ocr_box_to_pdf_rect(
            transformed_box,
            fitz.Rect(0, 0, 400, 500),
            900,
            1100,
            transform_matrix=transform,
            source_width=800,
            source_height=1000,
        )
        self.assertAlmostEqual(pdf_rect.x0, 80.0, places=5)
        self.assertAlmostEqual(pdf_rect.y0, 100.0, places=5)
        self.assertAlmostEqual(pdf_rect.x1, 200.0, places=5)
        self.assertAlmostEqual(pdf_rect.y1, 150.0, places=5)


if __name__ == "__main__":
    unittest.main()
