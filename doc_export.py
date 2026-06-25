# -*- coding: utf-8 -*-
"""
doc_export.py
==============
세무 자문 결과(마크다운 텍스트)를 docx / pdf 파일로 변환하는 모듈입니다.

- md  : 별도 변환 없이 그대로 .md 파일로 저장 (tax_advisor_engine.py의 save_response와 동일)
- docx: python-docx 사용. 마크다운 헤더(###)와 굵게(**) 표시를 기본적인 Word 서식으로 변환.
- pdf : fpdf2 사용. 한글 출력을 위해 Windows 시스템 폰트(맑은 고딕)를 자동으로 찾아 사용.
        폰트를 찾지 못하면 PDF 생성을 건너뛰고 명확한 안내 메시지를 반환.

Windows 환경 기준으로 작성됨.
"""

import re
from pathlib import Path
from datetime import datetime


# ----------------------------------------------------------------------
# 공통: 마크다운 본문을 줄 단위로 단순 파싱
# ----------------------------------------------------------------------
def _parse_markdown_lines(md_text: str):
    """
    마크다운 텍스트를 (종류, 내용) 튜플 리스트로 변환.
    종류: 'h3', 'h2', 'h1', 'bullet', 'text', 'blank'
    굵게(**...**) 표시는 내용 안에 마커를 남겨 호출부에서 처리.
    """
    lines = []
    for raw_line in md_text.split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            lines.append(("blank", ""))
        elif line.startswith("### "):
            lines.append(("h3", line[4:].strip()))
        elif line.startswith("## "):
            lines.append(("h2", line[3:].strip()))
        elif line.startswith("# "):
            lines.append(("h1", line[2:].strip()))
        elif line.strip().startswith("- "):
            lines.append(("bullet", line.strip()[2:].strip()))
        else:
            lines.append(("text", line.strip()))
    return lines


def _split_bold(text: str):
    """**굵게** 표시를 (텍스트, 굵음여부) 조각 리스트로 분리"""
    parts = []
    pattern = re.compile(r"\*\*(.+?)\*\*")
    last_end = 0
    for m in pattern.finditer(text):
        if m.start() > last_end:
            parts.append((text[last_end:m.start()], False))
        parts.append((m.group(1), True))
        last_end = m.end()
    if last_end < len(text):
        parts.append((text[last_end:], False))
    if not parts:
        parts = [(text, False)]
    return parts


# ----------------------------------------------------------------------
# DOCX 변환
# ----------------------------------------------------------------------
def build_docx_bytes(title: str, question: str, response_md: str) -> bytes:
    """
    질의/회신을 Word 문서(.docx) 형식의 바이트로 변환 (디스크에 쓰지 않음).

    설계 의도 (2026-06-25 — 웹 배포 지원):
    - 기존에는 output_path(파일 경로)를 받아 디스크에 직접 .save()하는 방식이었음.
    - PC 로컬 환경에서는 괜찮지만, 웹 서버(Streamlit Cloud)에는 사용자가 지정한
      로컬 경로(예: D드라이브 경로)가 존재하지 않아 디스크 쓰기 자체가 실패함.
    - 그래서 디스크에 쓰는 대신, 메모리(BytesIO)에 담아 바이트로 반환하도록 변경.
      이렇게 하면 호출하는 쪽(streamlit_ui.py)에서 st.download_button으로 바로
      사용자 컴퓨터에 다운로드시킬 수 있고, 동시에 로컬 환경에서도 그 바이트를
      파일로 저장하면 똑같이 동작함 (한 번 만들고 양쪽에 다 쓸 수 있는 구조).

    Parameters
    ----------
    title : str
        문서 제목
    question : str
        사용자 질문 (또는 종합 문서의 경우 안내문)
    response_md : str
        마크다운 형식의 AI 회신 본문

    Returns
    -------
    bytes
        .docx 파일의 바이트 내용
    """
    from io import BytesIO
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # 기본 폰트 설정 (한글 가독성을 위해 맑은 고딕 우선)
    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style.font.size = Pt(10.5)

    # 제목
    title_p = doc.add_heading(title, level=1)
    title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 메타 정보
    meta = doc.add_paragraph()
    meta.add_run(f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").italic = True

    if question:
        doc.add_heading("질의 내용", level=2)
        doc.add_paragraph(question)

    doc.add_heading("AI 회신", level=2)

    for kind, content in _parse_markdown_lines(response_md):
        if kind == "blank":
            continue
        elif kind == "h3":
            doc.add_heading(content, level=3)
        elif kind == "h2":
            doc.add_heading(content, level=2)
        elif kind == "h1":
            doc.add_heading(content, level=1)
        elif kind == "bullet":
            p = doc.add_paragraph(style="List Bullet")
            for text, is_bold in _split_bold(content):
                run = p.add_run(text)
                run.bold = is_bold
        else:  # text
            p = doc.add_paragraph()
            for text, is_bold in _split_bold(content):
                run = p.add_run(text)
                run.bold = is_bold

    footer = doc.add_paragraph()
    footer_run = footer.add_run("포스원 회계법인 세무질의 AI 시스템 자동 생성")
    footer_run.italic = True
    footer_run.font.size = Pt(9)
    footer_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ----------------------------------------------------------------------
# PDF 변환
# ----------------------------------------------------------------------
def _find_korean_font() -> str:
    """
    시스템에서 한글 폰트를 자동으로 탐색.
    Windows(로컬 PC)와 Linux(웹 서버) 양쪽 환경을 모두 고려함.

    주의: .ttc(TrueType Collection, 여러 폰트가 한 파일에 묶인 형식)는 fpdf2가
    제대로 처리하지 못해 렌더링 에러가 나는 경우가 있어 후보에서 제외함.
    .ttf(단일 폰트 파일) 형식만 사용함.

    찾지 못하면 빈 문자열 반환 (이 경우 PDF 생성은 건너뛰고 md/docx로 안내).
    """
    candidates = [
        # Windows (로컬 PC)
        r"C:\Windows\Fonts\malgun.ttf",       # 맑은 고딕
        r"C:\Windows\Fonts\malgunbd.ttf",     # 맑은 고딕 Bold
        # Linux (Streamlit Cloud 등 웹 서버 — 나눔고딕이 설치되어 있는 경우)
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return ""


def build_pdf_bytes(title: str, question: str, response_md: str) -> bytes:
    """
    질의/회신을 PDF 형식의 바이트로 변환 (디스크에 쓰지 않음).
    한글 출력을 위해 시스템 한글 폰트를 사용함(없으면 RuntimeError).

    Parameters
    ----------
    title, question, response_md : build_docx_bytes와 동일

    Returns
    -------
    bytes
        PDF 파일의 바이트 내용

    Raises
    ------
    RuntimeError
        시스템에서 한글 폰트를 찾지 못한 경우 (이 환경에서는 PDF 생성 불가).
        웹 배포 환경(Streamlit Cloud 기본 이미지)에는 한글 폰트가 없는 경우가
        많으므로, 이 경우 호출하는 쪽에서 PDF를 건너뛰고 md/docx로 안내해야 함.
    """
    from fpdf import FPDF

    font_path = _find_korean_font()
    if not font_path:
        raise RuntimeError(
            "한글 폰트를 찾을 수 없어 PDF를 생성할 수 없습니다.\n"
            "로컬 PC: Windows 폰트 폴더(C:\\Windows\\Fonts)에 malgun.ttf가 있는지 확인하세요.\n"
            "웹 배포 환경: 서버에 한글 폰트가 설치되어 있지 않습니다.\n"
            "대신 Word(.docx) 또는 마크다운(.md) 형식으로 받아주세요."
        )

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("korean", "", font_path)
    pdf.set_font("korean", size=14)

    pdf.multi_cell(0, 10, title)
    pdf.set_font("korean", size=9)
    pdf.multi_cell(0, 6, f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    pdf.ln(4)

    if question:
        pdf.set_font("korean", size=12)
        pdf.multi_cell(0, 8, "질의 내용")
        pdf.set_font("korean", size=10)
        pdf.multi_cell(0, 6, question)
        pdf.ln(4)

    pdf.set_font("korean", size=12)
    pdf.multi_cell(0, 8, "AI 회신")
    pdf.ln(1)

    for kind, content in _parse_markdown_lines(response_md):
        if kind == "blank":
            pdf.ln(2)
            continue
        plain_text = content.replace("**", "")  # PDF에서는 굵게 폰트 별도 미적용(가독성 우선 단순화)
        if kind in ("h1", "h2", "h3"):
            pdf.set_font("korean", size=12)
            pdf.ln(2)
            pdf.multi_cell(0, 7, plain_text)
            pdf.set_font("korean", size=10)
        elif kind == "bullet":
            pdf.multi_cell(0, 6, f"  - {plain_text}")
        else:
            pdf.multi_cell(0, 6, plain_text)

    pdf.ln(6)
    pdf.set_font("korean", size=8)
    pdf.multi_cell(0, 5, "포스원 회계법인 세무질의 AI 시스템 자동 생성")

    return bytes(pdf.output())


# ----------------------------------------------------------------------
# 통합 함수 (메모리 바이트 반환 — 웹 다운로드 버튼 / 로컬 파일 저장 양쪽에서 재사용)
# ----------------------------------------------------------------------
def build_export_bytes(
    title: str,
    question: str,
    response_md: str,
    formats: list = None,
) -> dict:
    """
    여러 형식의 파일을 메모리(바이트)로 한 번에 생성.

    설계 의도 (2026-06-25 — 웹 배포 지원):
    - 기존 export_response()는 디스크에 직접 저장하는 방식만 지원했음. 이 함수는
      디스크에 쓰지 않고 바이트만 만들어 반환하므로, 호출하는 쪽에서
      st.download_button(웹에서 사용자 컴퓨터로 다운로드)이나, 원한다면 직접
      파일로 저장(로컬 PC)하는 데 그대로 재사용할 수 있음.
    - md는 항상 만들 수 있음. docx도 항상 만들 수 있음(폰트 의존성 없음).
      pdf는 시스템에 한글 폰트가 있어야만 만들 수 있으므로, 실패 시
      results["errors"]에 사유가 남고 results["pdf"]는 None으로 유지됨
      (다른 형식 생성에는 영향 없음).

    Returns
    -------
    dict
        {"md": bytes 또는 None, "docx": bytes 또는 None, "pdf": bytes 또는 None,
         "errors": [str, ...]}
    """
    formats = formats or ["md"]
    results = {"md": None, "docx": None, "pdf": None, "errors": []}

    if "md" in formats:
        try:
            content = (
                f"# {title}\n\n**생성일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                + (f"## 질의 내용\n{question}\n\n---\n\n" if question else "")
                + f"## AI 회신\n\n{response_md}\n\n---\n*포스원 회계법인 세무질의 AI 시스템 자동 생성*\n"
            )
            results["md"] = content.encode("utf-8")
        except Exception as e:
            results["errors"].append(f"md 생성 실패: {e}")

    if "docx" in formats:
        try:
            results["docx"] = build_docx_bytes(title, question, response_md)
        except Exception as e:
            results["errors"].append(f"docx 생성 실패: {e}")

    if "pdf" in formats:
        try:
            results["pdf"] = build_pdf_bytes(title, question, response_md)
        except Exception as e:
            results["errors"].append(f"pdf 생성 실패: {e}")

    return results


def safe_filename(filename: str = None) -> str:
    """파일명에 쓸 수 없는 문자를 제거하고, 비어 있으면 타임스탬프 기반 기본값을 만듦."""
    if not filename:
        filename = f"자문_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return re.sub(r'[\\/:*?"<>|]', "_", filename).strip() or "자문"


# ----------------------------------------------------------------------
# (하위 호환용) 로컬 PC 디스크 저장 함수
# ----------------------------------------------------------------------
def export_response(
    title: str,
    question: str,
    response_md: str,
    output_dir: Path,
    filename: str = None,
    formats: list = None,
) -> dict:
    """
    여러 형식으로 한 번에 '디스크에' 저장하는 함수. PC 로컬 환경에서만 사용 권장.

    주의: 웹 배포 환경(Streamlit Cloud 등)에서는 output_dir이 실제로 쓸 수 있는
    경로가 아닐 가능성이 높음(사용자가 PC 기준으로 적어둔 경로가 서버에는 없음).
    웹 환경에서는 이 함수 대신 build_export_bytes()로 바이트를 만들어
    st.download_button으로 사용자에게 직접 내려주는 방식을 사용해야 함.

    Parameters
    ----------
    title : str
        문서 제목
    question : str
        질문 내용 (종합 문서의 경우 빈 문자열 가능)
    response_md : str
        마크다운 회신 본문
    output_dir : Path
        저장 폴더
    filename : str, optional
        파일명(확장자 제외). None이면 "자문_YYYYMMDD_HHMMSS" 자동 생성.
    formats : list[str], optional
        ["md", "docx", "pdf"] 중 선택. None이면 ["md"]만 저장.

    Returns
    -------
    dict
        {"md": Path 또는 None, "docx": ..., "pdf": ..., "errors": [str, ...]}
    """
    filename = safe_filename(filename)
    built = build_export_bytes(title, question, response_md, formats=formats)

    results = {"md": None, "docx": None, "pdf": None, "errors": list(built["errors"])}

    try:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        # 폴더 자체를 못 만들면(웹 서버의 잘못된 경로 등) 모든 디스크 저장이
        # 의미 없으므로, 명확한 에러만 남기고 즉시 반환 (예외를 던지지 않음 —
        # 호출하는 쪽 화면이 깨지지 않도록 함)
        results["errors"].append(f"저장 폴더를 만들 수 없습니다 ({output_dir}): {e}")
        return results

    ext_map = {"md": ".md", "docx": ".docx", "pdf": ".pdf"}
    for fmt, ext in ext_map.items():
        data = built.get(fmt)
        if data is None:
            continue
        try:
            file_path = output_dir / f"{filename}{ext}"
            file_path.write_bytes(data)
            results[fmt] = file_path
        except Exception as e:
            results["errors"].append(f"{fmt} 저장 실패: {e}")

    return results
