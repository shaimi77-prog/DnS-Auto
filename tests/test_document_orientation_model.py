"""공식 문서 방향 ONNX의 배포·추론·보정 방향 회귀시험."""

import hashlib
from pathlib import Path
import unittest

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

from page_preprocessing import (
    DocumentOrientationClassifier,
    OrientationPrediction,
    PagePreprocessConfig,
    PagePreprocessor,
)


SOURCE = Path(__file__).resolve().parents[1]
MODEL = SOURCE / "ocr_models" / "PP-LCNet_x1_0_doc_ori.onnx"
EXPECTED_SHA256 = "1db9914a3beb04181fde445b2fef96b850072f89a2fa8aa71ebef4ed03b8074f"


def document_image():
    image = Image.new("RGB", (900, 1200), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default(size=22)
    draw.rectangle((60, 60, 840, 1140), outline="black", width=3)
    draw.text((100, 100), "DOCUMENT ORIENTATION TEST", fill="black", font=font)
    for row in range(12):
        y = 190 + row * 70
        draw.text(
            (100, y),
            f"Row {row + 1:02d}  ABCDEFG  1234567890",
            fill="black",
            font=font,
        )
        draw.line((100, y + 35, 800, y + 35), fill="gray", width=2)
    return np.asarray(image)[:, :, ::-1].copy()

class DocumentOrientationModelTests(unittest.TestCase):
    def test_official_converted_model_hash_and_specs(self):
        self.assertTrue(MODEL.is_file())
        self.assertEqual(hashlib.sha256(MODEL.read_bytes()).hexdigest(), EXPECTED_SHA256)
        for spec_name in ("DnS_Auto.spec", "DnS_Auto_MCP.spec"):
            spec = (SOURCE / spec_name).read_text(encoding="utf-8")
            self.assertIn("ocr_models", spec)

    def test_model_classifies_four_document_directions(self):
        classifier = DocumentOrientationClassifier(MODEL)
        base = document_image()
        samples = {
            0: base,
            90: cv2.rotate(base, cv2.ROTATE_90_CLOCKWISE),
            180: cv2.rotate(base, cv2.ROTATE_180),
            270: cv2.rotate(base, cv2.ROTATE_90_COUNTERCLOCKWISE),
        }
        for expected, sample in samples.items():
            with self.subTest(expected=expected):
                self.assertEqual(classifier(sample).degrees, expected)

    def test_detected_clockwise_rotation_is_corrected_in_reverse(self):
        result = PagePreprocessor(
            PagePreprocessConfig(enabled=True, orientation_enabled=True),
            lambda _image: OrientationPrediction(90, 0.99, 0.80),
        ).preprocess(document_image())
        self.assertTrue(result.orientation_applied)
        self.assertEqual(result.detected_orientation, 90)
        self.assertEqual(result.orientation_correction, 270)


if __name__ == "__main__":
    unittest.main()