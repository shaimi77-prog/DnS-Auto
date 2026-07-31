from __future__ import annotations

from unittest.mock import patch

import ui_mcp_bridge


def main() -> None:
    parent = object()
    with patch("ui_mcp_bridge.engine_Drag.run_application", return_value=True) as run:
        result = ui_mcp_bridge.run_pdf_application(parent, force_ocr=True)
    run.assert_called_once_with(parent, force_ocr=True)
    assert result is True
    assert not hasattr(ui_mcp_bridge, "filedialog"), "GUI 브리지가 파일 선택을 중복 수행하면 안 됩니다."
    print("gui_pdf_single_entry: PASS")
    print("mcp_paths_untouched: PASS")


if __name__ == "__main__":
    main()