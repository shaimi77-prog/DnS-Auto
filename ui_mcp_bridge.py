"""Tkinter 메인 화면에서 기존 GUI 엔진으로 진입하는 얇은 브리지."""

import engine_Drag


def run_pdf_application(
    parent_root,
    force_ocr=False,
    pdf_collection_mode=engine_Drag.PDF_MODE_STANDARD,
):
    """PDF GUI 흐름을 한 번만 시작한다.

    양식·시트·헤더·프로필 선택은 engine_Drag.run_application 안에서 일관되게
    처리한다. MCP의 저장 프로필 및 대화형 실행은 mcp_server와
    interactive_runner의 별도 경로를 사용하므로 이 브리지의 영향을 받지 않는다.
    """
    if pdf_collection_mode == engine_Drag.PDF_MODE_STANDARD:
        return engine_Drag.run_application(parent_root, force_ocr=force_ocr)
    return engine_Drag.run_application(
        parent_root, force_ocr=force_ocr,
        pdf_collection_mode=pdf_collection_mode,
    )
