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
def export_to_docx(title: str, question: str, response_md: str, output_path: Path) -> Path:
    """
    질의/회신을 Word 문서(.docx)로 저장

    Parameters
    ----------
    title : str
        문서 제목
    question : str
        사용자 질문 (또는 종합 문서의 경우 안내문)
    response_md : str
        마크다운 형식의 AI 회신 본문
    output_path : Path
        저장할 .docx 파일 경로 (확장자 포함)

    Returns
    -------
    Path
        저장된 파일 경로
    """
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

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ----------------------------------------------------------------------
# PDF 변환
# ----------------------------------------------------------------------
def _find_korean_font() -> str:
    """
    Windows 시스템에서 한글 폰트(맑은 고딕 등)를 자동으로 탐색.
    찾지 못하면 빈 문자열 반환.
    """
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",       # 맑은 고딕
        r"C:\Windows\Fonts\malgunbd.ttf",     # 맑은 고딕 Bold
        r"C:\Windows\Fonts\gulim.ttc",        # 굴림 (대체)
        r"C:\Windows\Fonts\batang.ttc",       # 바탕 (대체)
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return ""


def export_to_pdf(title: str, question: str, response_md: str, output_path: Path) -> Path:
    """
    질의/회신을 PDF(.pdf)로 저장. 한글 출력을 위해 Windows 시스템 폰트를 사용함.

    Parameters
    ----------
    title, question, response_md, output_path : export_to_docx와 동일

    Returns
    -------
    Path
        저장된 파일 경로

    Raises
    ------
    FileNotFoundError
        시스템에서 한글 폰트를 찾지 못한 경우 (이 환경에서는 PDF 생성 불가)
    """
    from fpdf import FPDF

    font_path = _find_korean_font()
    if not font_path:
        raise FileNotFoundError(
            "한글 폰트(맑은 고딕)를 찾을 수 없어 PDF를 생성할 수 없습니다.\n"
            "Windows 시스템 폰트 폴더(C:\\Windows\\Fonts)에 malgun.ttf가 있는지 확인하세요.\n"
            "대신 Word(.docx) 또는 마크다운(.md) 형식으로 저장해주세요."
        )

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("malgun", "", font_path)
    pdf.set_font("malgun", size=14)

    pdf.multi_cell(0, 10, title)
    pdf.set_font("malgun", size=9)
    pdf.multi_cell(0, 6, f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    pdf.ln(4)

    if question:
        pdf.set_font("malgun", size=12)
        pdf.multi_cell(0, 8, "질의 내용")
        pdf.set_font("malgun", size=10)
        pdf.multi_cell(0, 6, question)
        pdf.ln(4)

    pdf.set_font("malgun", size=12)
    pdf.multi_cell(0, 8, "AI 회신")
    pdf.ln(1)

    for kind, content in _parse_markdown_lines(response_md):
        if kind == "blank":
            pdf.ln(2)
            continue
        plain_text = content.replace("**", "")  # PDF에서는 굵게 폰트 별도 미적용(가독성 우선 단순화)
        if kind in ("h1", "h2", "h3"):
            pdf.set_font("malgun", size=12)
            pdf.ln(2)
            pdf.multi_cell(0, 7, plain_text)
            pdf.set_font("malgun", size=10)
        elif kind == "bullet":
            pdf.multi_cell(0, 6, f"  - {plain_text}")
        else:
            pdf.multi_cell(0, 6, plain_text)

    pdf.ln(6)
    pdf.set_font("malgun", size=8)
    pdf.multi_cell(0, 5, "포스원 회계법인 세무질의 AI 시스템 자동 생성")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(output_path))
    return output_path


# ----------------------------------------------------------------------
# 통합 저장 함수
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
    여러 형식으로 한 번에 저장하는 통합 함수

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
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not filename:
        filename = f"자문_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    # 파일명에 쓸 수 없는 문자 제거
    filename = re.sub(r'[\\/:*?"<>|]', "_", filename).strip() or "자문"

    formats = formats or ["md"]
    results = {"md": None, "docx": None, "pdf": None, "errors": []}

    if "md" in formats:
        try:
            md_path = output_dir / f"{filename}.md"
            content = (
                f"# {title}\n\n**생성일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                + (f"## 질의 내용\n{question}\n\n---\n\n" if question else "")
                + f"## AI 회신\n\n{response_md}\n\n---\n*포스원 회계법인 세무질의 AI 시스템 자동 생성*\n"
            )
            md_path.write_text(content, encoding="utf-8")
            results["md"] = md_path
        except Exception as e:
            results["errors"].append(f"md 저장 실패: {e}")

    if "docx" in formats:
        try:
            results["docx"] = export_to_docx(title, question, response_md, output_dir / f"{filename}.docx")
        except Exception as e:
            results["errors"].append(f"docx 저장 실패: {e}")

    if "pdf" in formats:
        try:
            results["pdf"] = export_to_pdf(title, question, response_md, output_dir / f"{filename}.pdf")
        except Exception as e:
            results["errors"].append(f"pdf 저장 실패: {e}")

    return results
