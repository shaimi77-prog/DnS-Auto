# DnS Auto v1.0.0 Release

첫 안정 공개 버전입니다. 실행 파일은 소스 브랜치에 커밋하지 않고 GitHub Release의 DnS_Auto_Portable.zip으로 제공합니다.

## 검증

- 전체 자동 테스트 35건 통과(환경 조건 3건 제외)
- ETA·OCR·XLS 회귀 테스트 10건 통과
- 기준 파일·기준 페이지 자동 추천 후 전체 파일·페이지 수동 탐색 지원
- OCR 기준 페이지 확정 안내 및 기준 페이지에서 한 번만 수행하는 영역·키워드 설정 유지
- 이미지 페이지별 90도 단위 방향 보정, 평면 스캔 미세 기울기 보정 및 기준 앵커 검증 추가
- 보정 검증 실패 시 해당 페이지만 안전하게 제외하고 감사 로그에 시도·성공·폐기 사유 기록
- GUI 창 제목 및 Windows 파일 버전 1.0.0 확인
- MCP serverInfo.version 1.0.0 확인
- 별도 폴더 압축 해제 후 GUI 실행 및 MCP Excel 취합 확인

## SHA-256

`
98C4D42A96791B4F6D4C6A18EAEE77866124754F1F86B0A45227E69FC9FBAA60  DnS_Auto_Portable.zip
`

## 설치

ZIP 전체를 압축 해제한 뒤 DnS Auto.exe를 실행하십시오. MCP 사용자는 같은 폴더의 MCP_GUIDE.html을 먼저 확인하십시오.
