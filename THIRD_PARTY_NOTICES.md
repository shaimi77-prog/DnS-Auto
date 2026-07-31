# 제3자 구성요소 고지

DnS Auto는 다음 오픈소스 라이브러리와 OCR 모델을 사용합니다. 이 문서는
제3자 구성요소의 저작권을 대체하지 않으며, 각 프로젝트의 원문 라이선스가
우선합니다.

## 주요 Python 패키지

| 구성요소 | 검증 버전 | 라이선스 | 프로젝트 |
|---|---:|---|---|
| PyMuPDF | 1.28.0 | GNU AGPL-3.0 또는 Artifex 상용 라이선스 | https://pymupdf.readthedocs.io/ |
| openpyxl | 3.1.5 | MIT | https://openpyxl.readthedocs.io/ |
| Pillow | 11.3.0 | HPND | https://python-pillow.org/ |
| pywin32 | 312 | PSF | https://github.com/mhammond/pywin32 |
| xlrd | 2.0.2 | BSD | https://xlrd.readthedocs.io/ |
| NumPy | 2.4.6 | BSD-3-Clause | https://numpy.org/ |
| RapidOCR | 3.9.2 | Apache-2.0(프로그램 코드) | https://github.com/RapidAI/RapidOCR |
| ONNX Runtime | 1.28.0 | MIT | https://onnxruntime.ai/ |
| PyInstaller | 6.21.0 | GPL-2.0-or-later 및 부트로더 예외 | https://pyinstaller.org/ |

정확한 설치 의존성 범위는 `requirements.txt`와
`requirements-build.txt`를 참고하십시오.

## 주요 간접 의존성

RapidOCR, ONNX Runtime 및 openpyxl을 설치하면 다음 패키지도 함께 설치될
수 있습니다. 공개 EXE에 포함되는 실제 구성은 빌드 환경과 패키지 버전에
따라 달라질 수 있습니다.

| 구성요소 | 검증 버전 | 라이선스 |
|---|---:|---|
| colorlog | 6.12.0 | MIT |
| OmegaConf | 2.3.1 | BSD |
| opencv-python | 5.0.0.93 | Apache-2.0 및 포함된 제3자 고지 |
| pyclipper | 1.4.0 | MIT |
| PyYAML | 6.0.3 | MIT |
| Requests | 2.34.2 | Apache-2.0 |
| Shapely | 2.1.2 | BSD-3-Clause |
| six | 1.17.0 | MIT |
| tqdm | 4.70.0 | MPL-2.0 및 MIT |
| flatbuffers | 25.12.19 | Apache-2.0 |
| packaging | 26.2 | Apache-2.0 또는 BSD-2-Clause |
| protobuf | 7.35.1 | BSD-3-Clause |
| et-xmlfile | 2.0.0 | MIT |

NumPy와 OpenCV 바이너리는 자체 제3자 구성요소 고지를 포함할 수 있습니다.
GitHub Release를 만들 때는 최종 빌드 환경에 설치된 배포 패키지의
`licenses` 디렉터리와 제3자 고지 파일을 함께 보존해야 합니다.

## OCR 모델

RapidOCR 프로젝트는 프로그램 코드와 OCR 모델의 저작권 주체가 다를 수
있음을 고지합니다. 아래 모델은 RapidOCR 3.9.2의 모델 목록을 통해 배포되는
PaddleOCR 계열 모델이며, 원 모델 저작권은 Baidu/PaddleOCR 측에 있습니다.

| 파일 | SHA-256 |
|---|---|
| `ch_PP-OCRv5_det_mobile.onnx` | `4D97C44A20D30A81AAD087D6A396B08F786C4635742AFC391F6621F5C6AE78AE` |
| `ch_ppocr_mobile_v2.0_cls_mobile.onnx` | `E47ACEDF663230F8863FF1AB0E64DD2D82B838FCEB5957146DAB185A89D6215C` |
| `korean_PP-OCRv5_rec_mobile.onnx` | `CD6E2EA50F6943CA7271EB8C56A877A5A90720B7047FE9C41A2E541A25773C9B` |

원본 배포 위치:

- https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv5/det/ch_PP-OCRv5_det_mobile.onnx
- https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv4/cls/ch_ppocr_mobile_v2.0_cls_mobile.onnx
- https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv5/rec/korean_PP-OCRv5_rec_mobile.onnx

관련 프로젝트:

- RapidOCR: https://github.com/RapidAI/RapidOCR
- PaddleOCR: https://github.com/PaddlePaddle/PaddleOCR

## 호환성 판단

DnS Auto 자체 소스와 배포 실행 파일을 AGPL-3.0 조건으로 제공하고, 해당
버전의 완전한 대응 소스와 빌드 절차를 함께 공개하는 현재 배포 방식에서는
위 주요 라이브러리의 라이선스와 명백한 충돌은 확인되지 않았습니다.

핵심 조건은 다음과 같습니다.

- PyMuPDF를 포함한 배포물은 AGPL-3.0 의무를 충족하거나 Artifex 상용
  라이선스를 별도로 확보해야 합니다.
- Apache-2.0, MIT, BSD, PSF, HPND/MIT-CMU 및 MPL-2.0 구성요소의 저작권
  고지와 라이선스 조건을 유지해야 합니다.
- RapidOCR 모델은 프로그램 코드와 저작권 주체가 다르므로 모델 출처와
  Baidu/PaddleOCR 저작권 고지를 유지해야 합니다.
- PyInstaller의 예외는 생성된 실행 파일의 라이선스를 제한하지 않지만,
  포함된 각 의존성의 라이선스는 별도로 준수해야 합니다.
- 의존성 또는 OCR 모델을 교체·업데이트하면 Release 전에 다시 검토해야
  합니다.

이 문서는 공개 준비를 위한 기술적 검토이며 법률 자문을 대신하지 않습니다.
