"""Apply the approved PDF-mode wording to the single-line HTML guides."""

from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[3]


def update(name, replacements, insertion):
    path = ROOT / name
    text = path.read_text(encoding="utf-8")
    for old, new in replacements:
        text = text.replace(old, new)
    marker = "</main></body></html>"
    if insertion not in text:
        text = text.replace(marker, insertion + marker)
    path.write_text(text, encoding="utf-8")


mode_section = """<h2>PDF 취합 모드</h2><table><tr><th>모드</th><th>동작</th></tr><tr><td>신속 — OCR 건너뛰기</td><td>PDF에 포함된 문자만 사용합니다. 매핑 필드 하나라도 추출되지 않으면 해당 페이지 행 전체를 제외하며 빈 행을 만들지 않습니다.</td></tr><tr><td>표준 — 기본값</td><td>PDF 문자를 우선 사용하고 이미지 문자는 현재 기본 OCR 정책으로 인식합니다.</td></tr><tr><td>신중 — 실험적</td><td>이미지 문자를 표준보다 조금 더 크게 분석합니다. 시간이 늘 수 있으며 정확도 향상을 보장하지 않습니다.</td></tr></table><p class="note">프로그램 시작 시 항상 표준 모드입니다. 모드는 PDF 취합에만 적용되고 프로필에 저장되지 않습니다. 신속 모드는 정상 빈 필드가 있는 양식도 행 전체를 제외할 수 있습니다.</p>"""

update(
    "GUI_QUICK_START.html",
    [
        ("<h3>OCR 선택</h3><p>이미지 PDF는 기본 설정에서도 자동 OCR 처리됩니다. 복사한 글자가 깨지거나 화면 내용과 텍스트가 다를 때만 <code>PDF 취합 시 강제 OCR 사용</code>을 선택합니다.</p>", "<h3>PDF 취합 모드</h3><p><b>표준</b>이 기본값입니다. OCR을 건너뛰려면 신속, 이미지 문자를 더 크게 분석하려면 신중(실험적)을 선택합니다.</p>"),
        ("<code>PDF 취합 시 강제 OCR 사용</code>을 선택해 다시 처리", "표준 모드로 다시 처리하거나 신중(실험적) 모드를 비교"),
    ],
    mode_section,
)

# Replace the obsolete force-OCR chapter as one bounded HTML section.
guide_path = ROOT / "GUI_GUIDE.html"
guide = guide_path.read_text(encoding="utf-8")
guide = guide.replace(
    "<td>PDF 취합 모드</td>\n<td>PDF 텍스트 레이어를 사용하지 않고 지정 영역을 OCR로 판독</td>",
    "<td>PDF 취합 모드</td>\n<td>신속·표준·신중 중 현재 실행의 PDF 문자·OCR 처리 정책을 선택</td>",
).replace(
    "스캔 PDF를\n처리할 때만 <code>PDF 취합 모드</code>을 선택한다.",
    "PDF 취합은 표준 모드가 기본이며, OCR을 생략할 때 신속, 이미지 문자를 더 크게 분석할 때 신중(실험적)을 선택한다.",
).replace(
    "<td><code>강제 OCR</code> 선택 후 <code>PDF 파일 취합(Drag)</code></td>",
    "<td><code>표준</code> 또는 <code>신중(실험적)</code> 선택 후 <code>PDF 파일 취합(Drag)</code></td>",
).replace(
    "<li>표준 또는 신중 모드에서는 텍스트 레이어를 사용하지 않고 지정 영역을 OCR로 판독</li>",
    "<li>신속은 OCR 없이 완전한 네이티브 행만 적재하고, 표준·신중은 필요한 이미지 문자를 OCR로 판독</li>",
).replace(
    "<li>텍스트가 깨지거나 화면 내용과 텍스트 레이어가 다를 때만 <code>PDF 취합 모드</code>을 선택한다.</li>",
    "<li>PDF 전용 패널에서 신속·표준·신중 중 현재 작업 정책을 선택한다. 기본값은 표준이다.</li>",
)
replacement = """<h3 id="73-pdf-취합-모드">7.3 PDF 취합 모드</h3>
<p><b>표준</b>은 기존 자동 처리와 같은 기본값이다. PDF 문자를 우선 사용하고 필요한 이미지 문자는 OCR로 인식한다.</p>
<p><b>신속</b>은 OCR을 전혀 실행하지 않는다. 매핑 필드 중 하나라도 PDF 문자로 얻지 못하면 페이지 행 전체를 제외하며 빈 행을 만들지 않는다. 정상 빈 필드가 있는 양식도 제외될 수 있다.</p>
<p><b>신중(실험적)</b>은 이미지 값 영역을 표준보다 조금 더 크게 분석한다. 처리시간이 늘 수 있고 정확도 향상을 보장하지 않는다.</p>
<p>모드는 PDF 취합에만 적용되고 프로필에 저장되지 않는다. 프로그램을 새로 시작하면 항상 표준이 선택된다.</p>
"""
guide = re.sub(
    r'<h3 id="73-[^"]+">.*?(?=<h3 id="74-처리-절차">)',
    replacement,
    guide,
    flags=re.S,
)
guide_path.write_text(guide, encoding="utf-8")

update(
    "USER_GUIDE.html",
    [],
    mode_section,
)

update(
    "GUI_GUIDE.html",
    [
        ("PDF 취합 시 강제 OCR 사용", "PDF 취합 모드"),
        ("강제 OCR 선택 시", "표준 또는 신중 모드에서는"),
        ("강제 OCR을 선택한다", "표준 또는 신중 모드를 선택한다"),
        ("강제 OCR을 선택하여", "표준 또는 신중 모드로"),
    ],
    mode_section,
)

update(
    "MCP_GUIDE.html",
    [],
    mode_section + "<p class=\"note\">MCP의 <code>pdf_collection_mode</code>은 <code>fast</code>, <code>standard</code>, <code>careful</code>을 허용하며 생략하면 <code>standard</code>입니다. 기존 <code>force_ocr</code> 입력은 하위 호환 진단용으로 유지되며 일반 GUI에는 표시되지 않습니다.</p>",
)
