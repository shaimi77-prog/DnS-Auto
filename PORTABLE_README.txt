DnS Auto 통합 포터블 빠른 시작
================================

이 폴더에는 두 실행 파일이 있습니다.

1. 사용자가 직접 화면에서 작업할 때
   - "DnS Auto.exe"를 더블클릭합니다.
   - 자세한 방법: GUI_GUIDE.html

2. AI 프로그램으로 작업할 때
   - "DnS Auto MCP.exe"를 직접 더블클릭하지 않습니다.
   - 사용하는 AI의 로컬 MCP 설정에서 이 EXE의 절대경로를 command로 한 번 등록합니다.
   - 자세한 방법: MCP_GUIDE.html

공유 폴더
---------
- inputs: AI로 처리할 양식과 원본 파일
- outputs: AI MCP 작업 결과
- profiles\sheet: 두 프로그램이 공유하는 Excel Sheet 프로필
- profiles\pdf: 두 프로그램이 공유하는 PDF 영역 매핑 프로필

권장 사용 흐름
--------------
1. 처음 사용하는 양식은 DnS Auto.exe 또는 MCP 대화형 설정으로 헤더·영역을 지정합니다.
2. 설정을 프로필로 저장합니다.
3. 다음부터 AI에 프로필명을 지정해 무인 취합합니다.

주의
----
- EXE만 따로 옮기지 말고 DnS Auto 폴더 전체를 유지하세요.
- 폴더를 옮기면 AI 프로그램의 MCP command 경로도 새 위치로 수정하세요.
- DOCX, HWP/HWPX 또는 XLS 문서 변환은 해당 프로그램이 PC에 설치되어 있어야 합니다.