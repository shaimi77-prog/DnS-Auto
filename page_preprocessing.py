"""페이지별 OCR 이미지 방향·기울기 보정 공용 코어.

네이티브 PDF 좌표나 프로파일 저장 형식은 변경하지 않는다. 이 모듈이
만드는 이미지와 변환 행렬은 한 페이지의 OCR 호출 동안에만 유효하다.
"""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Mapping

import cv2
import numpy as np


@dataclass(frozen=True)
class PagePreprocessConfig:
    enabled: bool = False
    orientation_enabled: bool = False
    deskew_enabled: bool = False
    orientation_min_confidence: float = 0.92
    orientation_min_margin: float = 0.15
    deskew_min_angle: float = 0.3
    deskew_max_angle: float = 10.0
    deskew_min_lines: int = 5
    deskew_max_mad: float = 1.5


@dataclass(frozen=True)
class OrientationPrediction:
    degrees: int
    confidence: float
    margin: float
    probabilities: Mapping[int, float] = field(default_factory=dict)


@dataclass(frozen=True)
class SkewEstimate:
    angle: float = 0.0
    confidence: float = 0.0
    valid_line_count: int = 0
    angle_mad: float = 0.0


@dataclass
class PagePreprocessResult:
    status: str
    processed_image: np.ndarray
    pdf_rotation: int = 0
    detected_orientation: int = 0
    orientation_correction: int = 0
    orientation_confidence: float = 0.0
    orientation_margin: float = 0.0
    orientation_applied: bool = False
    skew_angle: float = 0.0
    skew_confidence: float = 0.0
    valid_line_count: int = 0
    angle_dispersion: float = 0.0
    deskew_applied: bool = False
    reference_validation: str = "not_requested"
    failure_reason: str | None = None
    transform_matrix: np.ndarray = field(
        default_factory=lambda: np.eye(3, dtype=np.float64)
    )


OrientationClassifier = Callable[[np.ndarray], OrientationPrediction]
ReferenceValidator = Callable[[PagePreprocessResult], bool | tuple[bool, str]]


def orientation_prediction(probabilities: Mapping[int, float]) -> OrientationPrediction:
    """확률표를 최고 클래스와 1·2위 차이로 정규화한다."""
    normalized = {
        int(degrees) % 360: float(score)
        for degrees, score in probabilities.items()
        if int(degrees) % 360 in (0, 90, 180, 270)
    }
    if not normalized:
        return OrientationPrediction(0, 0.0, 0.0, {})
    ranked = sorted(normalized.items(), key=lambda item: item[1], reverse=True)
    degrees, confidence = ranked[0]
    runner_up = ranked[1][1] if len(ranked) > 1 else 0.0
    return OrientationPrediction(
        degrees,
        confidence,
        confidence - runner_up,
        normalized,
    )

class DocumentOrientationClassifier:
    """공식 PP-LCNet 문서 방향 ONNX를 필요할 때 한 번만 연다."""

    labels = (0, 90, 180, 270)

    def __init__(self, model_path: str | Path | None = None):
        self.model_path = Path(model_path) if model_path else (
            Path(__file__).resolve().parent
            / "ocr_models"
            / "PP-LCNet_x1_0_doc_ori.onnx"
        )
        self._session = None
        self._input_name = None

    def _ensure_session(self):
        if self._session is not None:
            return
        import onnxruntime as ort

        if not self.model_path.is_file():
            raise FileNotFoundError(f"문서 방향 모델이 없습니다: {self.model_path}")
        self._session = ort.InferenceSession(
            str(self.model_path), providers=["CPUExecutionProvider"]
        )
        self._input_name = self._session.get_inputs()[0].name

    @staticmethod
    def _preprocess(image: np.ndarray) -> np.ndarray:
        image = _as_bgr(image)
        height, width = image.shape[:2]
        scale = 256.0 / min(width, height)
        resized = cv2.resize(
            image,
            (round(width * scale), round(height * scale)),
            interpolation=cv2.INTER_LINEAR,
        )
        top = (resized.shape[0] - 224) // 2
        left = (resized.shape[1] - 224) // 2
        cropped = resized[top : top + 224, left : left + 224]
        tensor = cropped.astype(np.float32) / 255.0
        tensor = (
            tensor - np.asarray([0.485, 0.456, 0.406], dtype=np.float32)
        ) / np.asarray([0.229, 0.224, 0.225], dtype=np.float32)
        return np.transpose(tensor, (2, 0, 1))[None, ...]

    def __call__(self, image: np.ndarray) -> OrientationPrediction:
        self._ensure_session()
        scores = np.asarray(
            self._session.run(None, {self._input_name: self._preprocess(image)})[0]
        ).reshape(-1)
        return orientation_prediction(dict(zip(self.labels, scores.tolist())))

def _as_bgr(image: np.ndarray) -> np.ndarray:
    array = np.asarray(image)
    if array.ndim == 2:
        return cv2.cvtColor(array, cv2.COLOR_GRAY2BGR)
    if array.ndim != 3:
        raise ValueError("페이지 이미지는 2차원 또는 3차원 배열이어야 합니다.")
    if array.shape[2] == 4:
        return cv2.cvtColor(array, cv2.COLOR_BGRA2BGR)
    if array.shape[2] != 3:
        raise ValueError("지원하지 않는 페이지 이미지 채널 수입니다.")
    return np.ascontiguousarray(array)


def _right_angle_transform(width: int, height: int, degrees: int) -> tuple[np.ndarray, tuple[int, int]]:
    degrees %= 360
    if degrees == 0:
        return np.eye(3, dtype=np.float64), (width, height)
    if degrees == 90:
        return np.array([[0, -1, height - 1], [1, 0, 0], [0, 0, 1]], dtype=np.float64), (height, width)
    if degrees == 180:
        return np.array([[-1, 0, width - 1], [0, -1, height - 1], [0, 0, 1]], dtype=np.float64), (width, height)
    if degrees == 270:
        return np.array([[0, 1, 0], [-1, 0, width - 1], [0, 0, 1]], dtype=np.float64), (height, width)
    raise ValueError(f"90도 단위가 아닌 방향입니다: {degrees}")


def rotate_right_angle(image: np.ndarray, degrees: int) -> tuple[np.ndarray, np.ndarray]:
    """이미지를 시계 방향으로 회전하고 원본→결과 3x3 행렬을 반환한다."""
    image = _as_bgr(image)
    degrees %= 360
    transform, _size = _right_angle_transform(image.shape[1], image.shape[0], degrees)
    if degrees == 0:
        return image.copy(), transform
    rotate_code = {
        90: cv2.ROTATE_90_CLOCKWISE,
        180: cv2.ROTATE_180,
        270: cv2.ROTATE_90_COUNTERCLOCKWISE,
    }[degrees]
    return cv2.rotate(image, rotate_code), transform


def estimate_skew(image: np.ndarray, max_angle: float = 15.0) -> SkewEstimate:
    """여러 텍스트 계열 선분의 중앙값과 MAD로 평면 문서 기울기를 추정한다."""
    image = _as_bgr(image)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    if min(gray.shape[:2]) < 32:
        return SkewEstimate()
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.threshold(
        blurred, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU
    )[1]
    kernel_width = max(15, gray.shape[1] // 40)
    connected = cv2.morphologyEx(
        binary,
        cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_width, 1)),
    )
    edges = cv2.Canny(connected, 50, 150, apertureSize=3)
    minimum_length = max(30, gray.shape[1] // 12)
    lines = cv2.HoughLinesP(
        edges,
        1,
        np.pi / 1800,
        threshold=max(20, minimum_length // 2),
        minLineLength=minimum_length,
        maxLineGap=max(8, gray.shape[1] // 100),
    )
    if lines is None:
        return SkewEstimate()
    angles = []
    weights = []
    for x1, y1, x2, y2 in np.asarray(lines).reshape(-1, 4):
        dx = float(x2 - x1)
        dy = float(y2 - y1)
        if abs(dx) < 1:
            continue
        angle = float(np.degrees(np.arctan2(dy, dx)))
        if angle > 90:
            angle -= 180
        elif angle < -90:
            angle += 180
        if abs(angle) > max_angle:
            continue
        length = float(np.hypot(dx, dy))
        angles.append(angle)
        weights.append(length)
    if not angles:
        return SkewEstimate()
    angle_array = np.asarray(angles, dtype=np.float64)
    weight_array = np.asarray(weights, dtype=np.float64)
    order = np.argsort(angle_array)
    ordered_angles = angle_array[order]
    ordered_weights = weight_array[order]
    midpoint = ordered_weights.sum() / 2
    median_index = int(np.searchsorted(np.cumsum(ordered_weights), midpoint))
    median = float(ordered_angles[min(median_index, len(ordered_angles) - 1)])
    mad = float(np.median(np.abs(angle_array - median)))
    consistency = max(0.0, 1.0 - mad / 3.0)
    evidence = min(1.0, len(angles) / 12.0)
    return SkewEstimate(
        angle=median,
        confidence=consistency * evidence,
        valid_line_count=len(angles),
        angle_mad=mad,
    )


def rotate_fine(image: np.ndarray, correction_degrees: float) -> tuple[np.ndarray, np.ndarray]:
    """잘림을 막도록 캔버스를 확장해 미세 회전한다."""
    image = _as_bgr(image)
    height, width = image.shape[:2]
    center = (width / 2.0, height / 2.0)
    matrix = cv2.getRotationMatrix2D(center, correction_degrees, 1.0)
    cosine = abs(matrix[0, 0])
    sine = abs(matrix[0, 1])
    new_width = int(np.ceil(height * sine + width * cosine))
    new_height = int(np.ceil(height * cosine + width * sine))
    matrix[0, 2] += new_width / 2.0 - center[0]
    matrix[1, 2] += new_height / 2.0 - center[1]
    rotated = cv2.warpAffine(
        image,
        matrix,
        (new_width, new_height),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(255, 255, 255),
    )
    homogeneous = np.vstack([matrix, [0.0, 0.0, 1.0]])
    return rotated, homogeneous


def inverse_transform_points(
    points: np.ndarray,
    transform_matrix: np.ndarray,
) -> np.ndarray:
    """결과 이미지 좌표를 전처리 전 이미지 좌표로 되돌린다."""
    coordinates = np.asarray(points, dtype=np.float64).reshape(-1, 2)
    homogeneous = np.column_stack(
        [coordinates, np.ones(len(coordinates), dtype=np.float64)]
    )
    inverse = np.linalg.inv(np.asarray(transform_matrix, dtype=np.float64))
    restored = (inverse @ homogeneous.T).T
    restored[:, 0] /= restored[:, 2]
    restored[:, 1] /= restored[:, 2]
    return restored[:, :2]


def inverse_transform_box(
    box: np.ndarray,
    transform_matrix: np.ndarray,
) -> np.ndarray:
    """OCR 사각형 네 점을 원래 페이지 이미지 좌표계로 복원한다."""
    return inverse_transform_points(np.asarray(box), transform_matrix)

class PagePreprocessor:
    """모델 세션만 재사용하고 모든 판정·이미지는 호출별로 격리한다."""

    def __init__(
        self,
        config: PagePreprocessConfig | None = None,
        orientation_classifier: OrientationClassifier | None = None,
    ):
        self.config = config or PagePreprocessConfig()
        self.orientation_classifier = orientation_classifier

    def preprocess(
        self,
        image: np.ndarray,
        *,
        pdf_rotation: int = 0,
        reference_validator: ReferenceValidator | None = None,
    ) -> PagePreprocessResult:
        current = _as_bgr(image).copy()
        result = PagePreprocessResult(
            status="unchanged",
            processed_image=current,
            pdf_rotation=int(pdf_rotation) % 360,
        )
        if not self.config.enabled:
            return result
        applied = False
        if self.config.orientation_enabled and self.orientation_classifier is not None:
            prediction = self.orientation_classifier(current)
            result.detected_orientation = int(prediction.degrees) % 360
            result.orientation_confidence = float(prediction.confidence)
            result.orientation_margin = float(prediction.margin)
            if (
                result.detected_orientation
                and result.orientation_confidence >= self.config.orientation_min_confidence
                and result.orientation_margin >= self.config.orientation_min_margin
            ):
                # 모델 클래스는 입력 문서가 시계 방향으로 돌아간 각도다.
                # 원상 복구에는 그 반대 각도를 적용한다.
                result.orientation_correction = (
                    360 - result.detected_orientation
                ) % 360
                current, transform = rotate_right_angle(
                    current, result.orientation_correction
                )
                result.transform_matrix = transform @ result.transform_matrix
                result.orientation_applied = True
                applied = True
        if self.config.deskew_enabled:
            skew = estimate_skew(current)
            result.skew_angle = skew.angle
            result.skew_confidence = skew.confidence
            result.valid_line_count = skew.valid_line_count
            result.angle_dispersion = skew.angle_mad
            if (
                self.config.deskew_min_angle <= abs(skew.angle) <= self.config.deskew_max_angle
                and skew.valid_line_count >= self.config.deskew_min_lines
                and skew.angle_mad <= self.config.deskew_max_mad
            ):
                current, transform = rotate_fine(current, skew.angle)
                result.transform_matrix = transform @ result.transform_matrix
                result.deskew_applied = True
                applied = True
        result.processed_image = current
        result.status = "corrected" if applied else "unchanged"
        if reference_validator is not None:
            validation = reference_validator(result)
            accepted, reason = (
                validation if isinstance(validation, tuple) else (bool(validation), "")
            )
            result.reference_validation = "accepted" if accepted else "rejected"
            if not accepted:
                result.status = "rejected"
                result.failure_reason = reason or "reference_mismatch"
        return result
