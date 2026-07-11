# -*- coding: utf-8 -*-
"""
streamlit_ui.py
================
포스원 회계법인 세무질의 AI 시스템 - 웹 UI (Streamlit)

기능:
- 자연어 세무 질의 입력 및 AI 회신 표시
- "대화 묶음(스레드)" 구조: 같은 묶음 안에서는 AI가 이전 질문/답변을 실제로 참고하여
  꼬리질문에 맥락 있게 답변함
- "새 주제 시작" 버튼: 현재 화면의 묶음을 백로그로 보내고 화면을 깨끗하게 비움
  (백로그는 삭제되지 않고 "이전 대화 보기"에서 다시 펼쳐볼 수 있음)
- 결과 저장: 기본은 타임스탬프 자동저장, 파일명/형식(md/docx/pdf) 직접 지정 가능
- "전체 종합" 버튼: 현재 묶음 또는 백로그 포함 전체를 AI가 하나의 완성된 자문 문서로 재구성
- 로그인 시 관리자(회계사)/직원 비밀번호로 역할이 자동 결정됨(역할 직접 선택 단계 없음).
  관리자만 PIN 관리, 지식베이스 확정 저장, 검증대기 불러오기에 접근 가능

실행 방법:
    streamlit run streamlit_ui.py

필요 환경변수(.env 또는 Streamlit Secrets):
    ADMIN_PASSWORD   관리자(회계사) 로그인 비밀번호 (기존 APP_PASSWORD도 호환됨)
    STAFF_PASSWORD   직원 로그인 비밀번호

Windows 환경 기준으로 작성됨.
"""

import os
import json
import re
import uuid
from pathlib import Path
from datetime import datetime

import streamlit as st

# .env 파일을 여기서 직접 한 번 읽어둠.
# (tax_advisor_engine.py도 자체적으로 .env를 읽지만, 그건 엔진 객체가 만들어질 때
#  실행됨. 비밀번호 확인은 그보다 먼저 이뤄져야 하므로, 여기서 미리 읽어서
#  APP_PASSWORD를 os.getenv로 바로 가져올 수 있게 함.)
from dotenv import load_dotenv
load_dotenv()

from tax_advisor_engine import TaxAdvisorEngine
from doc_export import build_export_bytes, safe_filename
from sheet_logger import SheetLogger

try:
    from google.genai import types as genai_types
except ImportError:
    genai_types = None


# ----------------------------------------------------------------------
# 페이지 설정
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="포스원 세무 자문 AI",
    page_icon=None,
    layout="centered",
)

st.markdown(
    """
    <style>
    /* ------------------------------------------------------------------
       포스원 세무 자문 AI 시스템 - 디자인 테마 (네이비 + 골드)
       기존 위젯 동작(버튼, 폼, 인풋)은 그대로 두고 색/모양만 입힌 CSS.
       Streamlit 기본 마크업 구조에 의존하므로, 버전 업그레이드 시
       data-testid 셀렉터가 바뀌면 일부 효과가 사라질 수 있음(동작에는
       영향 없음 — 순수 스타일 레이어).
       ------------------------------------------------------------------ */

    :root {
        --pf-navy: #0c2340;
        --pf-navy-light: #15335c;
        --pf-gold: #e0b020;
        --pf-gold-text: #2b1d00;
        --pf-gold-strong: #8a6314;
        --pf-text-strong: #0c2340;
        --pf-text-muted: #6b7280;
        --pf-border: #e5e7eb;
        --pf-sidebar-label: #8fa6c4;
        --pf-sidebar-border: rgba(255, 255, 255, 0.15);
    }

    /* 이전 시도에서 header[data-testid="stHeader"]의 자식 요소를 통째로
       숨겼다가, Streamlit이 사이드바 펼치기/접기 토글 버튼도 같은 header
       안에 두고 있어서 토글 버튼까지 같이 사라지는 부작용이 있었음
       (사이드바 전체가 화면에서 안 보이게 됨).
       Streamlit 내부 DOM 구조(메뉴 띠와 토글 버튼을 분리하는 정확한
       선택자)는 버전마다 바뀔 수 있어 추측성 선택자로 다시 건드리는
       대신, header는 절대 건드리지 않고 .block-container의 위쪽
       여백만 넉넉하게 줘서 제목이 헤더 뒤로 가려지지 않게 함.
       (헤더를 숨기지 않으므로 Share/⭐/⋮ 메뉴 띠는 그대로 보이지만,
       사이드바 토글 버튼이 사라지는 부작용보다는 안전한 선택임) */
    .block-container { max-width: 880px; padding-top: 3.5rem; }

    /* 페이지 타이틀 영역 ------------------------------------------------ */
    .pf-header-row {
        display: flex; align-items: baseline; justify-content: space-between;
        gap: 12px; margin-bottom: 0.3rem;
    }
    .pf-header-title {
        font-size: 2.1rem !important; font-weight: 700 !important;
        color: var(--pf-text-strong) !important;
        margin: 0 !important; line-height: 1.3 !important;
        letter-spacing: -0.01em !important;
    }
    .pf-header-caption {
        color: var(--pf-text-muted) !important; font-size: 0.95rem !important;
        margin: 0.2rem 0 0 0 !important;
    }
    .pf-header-accent {
        width: 44px !important; height: 3px !important;
        background-color: var(--pf-gold) !important;
        border-radius: 0 !important; margin: 0.55rem 0 0.1rem !important;
    }

    /* 질의/답변 카드 ----------------------------------------------------- */
    .qa-question {
        font-weight: 700; color: var(--pf-text-strong);
        margin-bottom: 0.35rem; font-size: 1.02rem;
    }
    .qa-meta {
        color: var(--pf-text-muted); font-size: 0.78rem;
        margin-bottom: 0.15rem;
    }
    /* "1. 질의 요지" 등 ### 헤딩이 답변 본문(markdown) 안에 ###으로 올 때 -
       Streamlit이 이를 h3로 렌더링하므로 본문 영역 안의 h3만 골드 강조 */
    div[data-testid="stMarkdownContainer"] h3 {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        color: var(--pf-gold-strong) !important;
        margin-top: 0.9rem !important;
        margin-bottom: 0.4rem !important;
    }

    /* 진행 중인 대화 배지 ------------------------------------------------- */
    .thread-badge {
        display: inline-flex; align-items: center; gap: 6px;
        background: #eef3f9; color: var(--pf-text-strong);
        border: 1px solid #d6e2ef;
        border-radius: 999px; padding: 5px 14px; font-size: 0.82rem;
        font-weight: 600;
    }

    /* st.container(border=True) 카드 -> 살짝 더 또렷하게 */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        border-color: var(--pf-border) !important;
    }

    /* ------------------------------------------------------------------
       질의 입력 카드(question_form_wrap) - 좌측 골드 세로 바로 포인트
       (로그인 카드의 상단 골드 줄과는 다른 패턴이지만, 같은 골드 색상을
       사용해 디자인 언어는 통일시킴)
       ------------------------------------------------------------------ */
    .st-key-question_form_wrap div[data-testid="stForm"] {
        background-color: #fdf8ec !important;
        border: 1px solid #ecdfb8 !important;
        border-left: 4px solid var(--pf-gold) !important;
        border-radius: 0 12px 12px 0 !important;
    }
    /* "세무 질의를 입력하세요" 라벨 볼드 처리
       (Streamlit이 라벨 텍스트를 <label> 바로 아래 두는 버전과, 그 안을
       <p>로 한 번 더 감싸는 버전이 있어 양쪽 다 대응) */
    .st-key-question_form_wrap div[data-testid="stForm"] label,
    .st-key-question_form_wrap div[data-testid="stForm"] label p {
        font-weight: 700 !important;
        color: var(--pf-text-strong) !important;
    }
    /* 질의 입력창(textarea) - 카드 배경(아이보리)에 색을 넣었으므로
       입력창 자체는 흰색으로 되돌려 "입력하는 곳"이 또렷하게 떠 보이도록.
       Streamlit이 textarea를 감싸는 baseweb 래퍼에도 기본 배경색을 칠하기
       때문에, 래퍼와 textarea 양쪽 모두 흰색으로 지정. */
    .st-key-question_form_wrap div[data-testid="stForm"] div[data-baseweb="textarea"],
    .st-key-question_form_wrap div[data-testid="stForm"] textarea {
        background-color: #ffffff !important;
        border-color: #e3d9bd !important;
    }
    .st-key-question_form_wrap div[data-testid="stForm"] textarea::placeholder {
        color: #a89968 !important;
    }
    .st-key-question_form_wrap div[data-testid="stForm"] textarea:focus {
        box-shadow: 0 0 0 2px var(--pf-gold) !important;
        border-color: var(--pf-gold) !important;
    }
    /* 조회 버튼 - 아이보리 카드 위에서 골드 버튼은 묻히므로 네이비로 채움.
       (전역 primary 버튼 골드 규칙보다 우선하도록 카드 범위로 한정) */
    .st-key-question_form_wrap div[data-testid="stForm"] button {
        background-color: var(--pf-navy) !important;
        border-color: var(--pf-navy) !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        letter-spacing: 0.02em !important;
    }
    .st-key-question_form_wrap div[data-testid="stForm"] button:hover {
        background-color: var(--pf-navy-light) !important;
        border-color: var(--pf-navy-light) !important;
    }
    .st-key-question_form_wrap div[data-testid="stForm"] button p {
        color: #ffffff !important;
    }

    /* ------------------------------------------------------------------
       로그인 패널 - 네이비 배경 (브라우저 호환성 문제 없는 순수 HTML 블록)
       st.container(key="pf_login_wrap")로 감싼 영역 안에서만 적용되도록
       .st-key-pf_login_wrap으로 범위를 한정함 (다른 st.form에 영향 없음).
       ------------------------------------------------------------------ */
    .pf-login-panel {
        background-color: var(--pf-navy);
        border-radius: 12px 12px 0 0;
        padding: 1.75rem 2rem 1.5rem;
        max-width: 460px;
        border-top: 4px solid var(--pf-gold);
        box-shadow: 0 0 0 1px rgba(224, 176, 32, 0.3);
    }
    .pf-login-eyebrow {
        font-size: 0.82rem; font-weight: 600; color: #cfe0f3;
        letter-spacing: 0.02em; margin: 0 0 4px;
    }
    .pf-login-heading {
        font-size: 1.2rem; font-weight: 600; color: var(--pf-gold);
        margin: 0 0 6px;
    }
    .pf-login-desc {
        font-size: 0.85rem; color: #d7e3f0; line-height: 1.55;
        margin: 0;
    }
    /* 로그인 패널 바로 아래의 form(입력창+버튼)만 같은 폭의 박스로 이어붙임 */
    .st-key-pf_login_wrap div[data-testid="stForm"] {
        max-width: 460px;
        background-color: var(--pf-navy-light);
        border: none !important;
        border-radius: 0 0 12px 12px;
        padding: 1.25rem 2rem 1.75rem !important;
        box-shadow: 0 0 0 1px rgba(224, 176, 32, 0.3);
    }
    .st-key-pf_login_wrap div[data-testid="stForm"] input {
        background-color: #f3f5f8 !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        color: var(--pf-navy) !important;
        caret-color: var(--pf-navy) !important;
    }
    .st-key-pf_login_wrap div[data-testid="stForm"] input::placeholder {
        color: #8a93a3 !important;
    }
    .st-key-pf_login_wrap div[data-testid="stForm"] input:focus {
        box-shadow: 0 0 0 2px var(--pf-gold) !important;
        border-color: var(--pf-gold) !important;
    }
    .st-key-pf_login_wrap button[kind="primaryFormSubmit"] {
        background-color: var(--pf-gold) !important;
        border-color: var(--pf-gold) !important;
        color: var(--pf-gold-text) !important;
    }

    /* primary 버튼(조회, 확정 저장 실행, 입장 등) -> 골드로 채움 */
    button[kind="primary"], button[kind="primaryFormSubmit"] {
        background-color: var(--pf-gold) !important;
        border-color: var(--pf-gold) !important;
        color: var(--pf-gold-text) !important;
        font-weight: 600 !important;
        letter-spacing: 0.02em !important;
    }
    button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
        background-color: #c79a1c !important;
        border-color: #c79a1c !important;
    }

    /* st.info / st.success / st.warning 박스 라운드 처리 */
    div[data-testid="stAlert"] {
        border-radius: 10px !important;
    }

    /* 구분선 여백 약간 축소 (섹션이 많아 답답해 보이는 것 방지) */
    hr { margin: 1.1rem 0; }

    /* ------------------------------------------------------------------
       사이드바 - 네이비 배경으로 메인 화면과 톤 통일
       ------------------------------------------------------------------ */
    section[data-testid="stSidebar"] {
        background-color: var(--pf-navy) !important;
        border-right: 2px solid var(--pf-gold) !important;
    }
    section[data-testid="stSidebar"] * {
        color: #ffffff !important;
        text-align: left !important;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #ffffff !important;
        font-weight: 600 !important;
        text-align: left !important;
    }
    /* 사이드바 보조 설명 텍스트(캡션, 작은 글씨)는 옅은 블루그레이로 */
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: var(--pf-sidebar-label) !important;
    }
    section[data-testid="stSidebar"] h3 {
        margin-top: 0.3rem;
    }
    /* 사이드바 구분선 */
    section[data-testid="stSidebar"] hr {
        border-color: var(--pf-sidebar-border) !important;
    }
    /* 사이드바 내 인라인 코드(백틱) - 경로 표시 등의 가독성 보정 */
    section[data-testid="stSidebar"] code {
        background-color: rgba(255, 255, 255, 0.12) !important;
        color: #e8f0fb !important;
        border-radius: 4px !important;
    }
    /* 사이드바 입력창/슬라이더/버튼 영역은 어두운 배경에 맞춰 대비 보정
       (이전에는 옅은 반투명 흰 배경 + 흰 글자라서, 배경이 거의 흰색으로
       보이는데 그 위에 흰 글자/마스킹 점이 써져 거의 안 보이는 문제가
       있었음 — 로그인 입력창과 동일한 원인. 배경을 또렷한 밝은 회백색,
       글자색을 네이비로 바꿔 항상 대비가 보장되도록 수정) */
    section[data-testid="stSidebar"] input[type="text"],
    section[data-testid="stSidebar"] input[type="password"],
    section[data-testid="stSidebar"] textarea {
        background-color: #f3f5f8 !important;
        border-color: rgba(255, 255, 255, 0.25) !important;
        color: var(--pf-navy) !important;
        caret-color: var(--pf-navy) !important;
    }
    section[data-testid="stSidebar"] input[type="text"]::placeholder,
    section[data-testid="stSidebar"] input[type="password"]::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {
        color: #8a93a3 !important;
    }
    section[data-testid="stSidebar"] button {
        background-color: rgba(255, 255, 255, 0.08) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        color: #ffffff !important;
        text-align: center !important;
    }
    section[data-testid="stSidebar"] button[kind="primary"] {
        background-color: var(--pf-gold) !important;
        border-color: var(--pf-gold) !important;
        color: var(--pf-gold-text) !important;
    }
    /* 사이드바 라디오 버튼 선택색 -> 골드 포인트 */
    section[data-testid="stSidebar"] [data-baseweb="radio"] [aria-checked="true"] > div:first-child {
        border-color: var(--pf-gold) !important;
        background-color: var(--pf-gold) !important;
    }
    /* 사이드바 슬라이더 트랙/손잡이 -> 골드 포인트 */
    section[data-testid="stSidebar"] [data-baseweb="slider"] div[role="slider"] {
        background-color: var(--pf-gold) !important;
    }
    /* 사이드바 안의 st.success("구글 시트 로깅 사용 중") 배지 - 어두운 배경에서도 또렷하게 */
    section[data-testid="stSidebar"] div[data-testid="stAlert"] {
        background-color: rgba(34, 197, 94, 0.15) !important;
        border: 1px solid rgba(34, 197, 94, 0.4) !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stAlert"] * {
        color: #b9f3cf !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="pf-header-row"><p class="pf-header-title">포스원 세무 자문 AI 시스템</p></div>'
    '<p class="pf-header-caption">기장 직원 / 회계사를 위한 세무질의 실시간 응답 도구</p>'
    '<div class="pf-header-accent"></div>',
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# 로그인 (관리자/직원 분리 — 2026-06-27 변경)
# ----------------------------------------------------------------------
# 설계 의도 (2026-06-27 변경 — 단일 공통 비밀번호에서 역할별 분리로 전환):
# - 기존에는 APP_PASSWORD 1개로 누구나 동일하게 들어온 뒤, 사이드바의 "사용자
#   구분" 라디오로 본인이 직접 "직원/회계사"를 선택했음. 이건 자기보고
#   (self-report)일 뿐 실제 권한 분리가 아니었음 — 직원도 라디오에서 "회계사"를
#   선택하면 관리자 기능(지식베이스 확정 저장)에 그대로 접근 가능했음.
# - 변경: 비밀번호 자체를 관리자용(ADMIN_PASSWORD)과 직원용(STAFF_PASSWORD)으로
#   분리. 로그인 화면에서 입력한 비밀번호 값으로 역할이 자동 결정되며, 이후
#   "역할을 직접 선택하는" 단계 자체가 없음 — 그래서 사이드바의 "사용자 구분"
#   라디오는 제거함(11번째 줄 아래 사이드바 섹션 참고).
# - 비밀번호 값은 코드에 직접 적지 않고 .env(로컬) 또는 Streamlit Secrets(웹
#   배포 시)의 ADMIN_PASSWORD / STAFF_PASSWORD 값으로 관리함.
# - 하위 호환: 기존 APP_PASSWORD만 설정된 환경(아직 이번 변경을 반영해 Secrets를
#   갱신하지 않은 경우)에서도 막히지 않도록, ADMIN_PASSWORD가 없으면
#   APP_PASSWORD를 관리자 비밀번호로 대신 사용함.
# - "관리자"라는 이름을 쓴 이유: "회계사"라는 직무명을 그대로 쓰면, 향후 다른
#   회계사가 추가로 이 시스템을 쓰게 될 때 "회계사=특정 한 사람"처럼 들려
#   어색해질 수 있음. "관리자"는 권한 레벨을 가리키는 말이라 여러 명으로
#   늘어나도 자연스러움 (실제로 관리자 역할을 쓰는 사람은 회계사임).
# - 향후 고객 로그인을 추가할 때는 이 구조에 "customer" 역할을 한 단(段) 더
#   추가하는 형태로 확장하면 됨 (개발계획서 8.6절 참고 — 고객 인증 방식은
#   별도로 설계 필요, 공통 비밀번호 방식이 아니라 개인별 식별이 필요함).
# ----------------------------------------------------------------------
# 로그인 (관리자/직원 분리, 비밀번호는 구글시트 "계정설정" 탭에서 관리)
# ----------------------------------------------------------------------
# 설계 의도 (2026-06-27 변경 — 단일 공통 비밀번호에서 역할별 분리로 전환):
# - 기존에는 APP_PASSWORD 1개로 누구나 동일하게 들어온 뒤, 사이드바의 "사용자
#   구분" 라디오로 본인이 직접 "직원/회계사"를 선택했음. 이건 자기보고
#   (self-report)일 뿐 실제 권한 분리가 아니었음 — 직원도 라디오에서 "회계사"를
#   선택하면 관리자 기능(지식베이스 확정 저장)에 그대로 접근 가능했음.
# - 변경: 비밀번호 자체를 관리자용/직원용으로 분리. 로그인 화면에서 입력한
#   비밀번호 값으로 역할이 자동 결정되며, 이후 "역할을 직접 선택하는" 단계
#   자체가 없음 — 그래서 사이드바의 "사용자 구분" 라디오는 제거함.
# - 비밀번호 저장 위치 (2026-06-27 추가 변경 — .env/Secrets 대신 구글시트):
#   PIN과 마찬가지로 평문이 아니라 sha256 해시만 구글시트 "계정설정" 탭에
#   저장함(SheetLogger.set_account_password 등). .env(로컬)나 Streamlit
#   Secrets(웹)에 미리 비밀번호를 적어둘 필요가 없어지고, 관리자가 로그인 후
#   사이드바에서 언제든 두 비밀번호를 직접 변경할 수 있음.
# - 최초 실행 시(계정설정 탭이 비어있는 경우): 로그인 화면 대신 "최초 계정
#   설정" 화면을 띄워 그 자리에서 관리자/직원 비밀번호를 처음 만들게 함. 이
#   화면은 인증 전 누구나 접근 가능한 상태이므로, 배포 직후 가능한 빨리
#   설정을 완료하는 것을 전제로 함(의도적으로 약한 추가 보호장치를 두지 않음
#   — 그것도 결국 별도 설정값이 필요해 ".env가 귀찮다"는 원래 문제를
#   되살리기 때문).
# - 구글시트 연동이 비활성 상태인 경우(.env/Secrets에 GOOGLE_SHEET_ID 등이
#   없는 경우)에는 하위 호환을 위해 기존 ADMIN_PASSWORD/STAFF_PASSWORD(또는
#   구버전 APP_PASSWORD) 환경변수 방식으로 자동 폴백함.
# - 향후 고객 로그인을 추가할 때는 이 구조에 "customer" 역할을 한 단(段) 더
#   추가하는 형태로 확장하면 됨 (개발계획서 8.6절 참고 — 고객 인증 방식은
#   별도로 설계 필요, 공통 비밀번호 방식이 아니라 개인별 식별이 필요함).
@st.cache_resource(show_spinner=False)
def get_login_sheet_logger():
    """
    로그인 단계에서만 쓰는 가벼운 SheetLogger 인스턴스.
    TaxAdvisorEngine은 로그인 통과 후에야 만들어지므로(엔진 초기화 자체가
    무거운 작업이라 비인증 사용자에게는 실행하지 않으려는 의도), 로그인
    체크 시점에는 이 별도 인스턴스로 "계정설정" 탭에 접근함. 이후
    TaxAdvisorEngine 내부에서 한 번 더 SheetLogger를 만들지만, 둘 다
    가벼운 클라이언트 객체일 뿐이라 중복 비용은 미미함.
    """
    return SheetLogger()


def render_first_time_account_setup(login_logger: SheetLogger):
    """
    '계정설정' 탭이 비어있을 때(앱을 처음 띄운 경우) 보여주는 화면.
    관리자/직원 비밀번호를 그 자리에서 처음 만들게 함.
    """
    login_wrap = st.container(key="pf_login_wrap")
    with login_wrap:
        st.markdown(
            '<div class="pf-login-panel">'
            '<p class="pf-login-eyebrow">최초 설정</p>'
            '<p class="pf-login-heading">관리자·직원 비밀번호를 만들어주세요</p>'
            '<p class="pf-login-desc">아직 비밀번호가 설정되지 않았습니다.<br>'
            '관리자(회계사)와 직원이 각각 사용할 비밀번호를 처음 만들어주세요.</p>'
            '</div>',
            unsafe_allow_html=True,
        )

        with st.form("first_time_setup_form"):
            admin_pw = st.text_input("관리자(회계사) 비밀번호 (4자 이상)", type="password")
            staff_pw = st.text_input("직원 비밀번호 (4자 이상)", type="password")
            submitted = st.form_submit_button("설정 완료", type="primary")

    if submitted:
        if len(admin_pw.strip()) < 4 or len(staff_pw.strip()) < 4:
            st.error("두 비밀번호 모두 4자 이상으로 설정해주세요.")
        elif admin_pw.strip() == staff_pw.strip():
            st.error("관리자와 직원 비밀번호는 서로 다르게 설정해주세요.")
        else:
            login_logger.set_account_password("admin", admin_pw)
            login_logger.set_account_password("staff", staff_pw)
            st.success("설정이 완료되었습니다. 이제 해당 비밀번호로 로그인해주세요.")
            st.rerun()


def _get_master_reset_password() -> str:
    """
    비상 복구 비밀번호 값. .env/Secrets의 MASTER_RESET_PASSWORD가 있으면 그 값을,
    없으면 기본값 "4119"를 사용함. 기본값을 그대로 두는 건 보안상 바람직하지
    않으므로, 이 값으로 로그인하면 화면에 경고 배너를 띄워 변경을 유도함.
    """
    return os.getenv("MASTER_RESET_PASSWORD", "").strip() or "4119"


def check_app_password():
    """
    입력한 비밀번호가 관리자/직원 비밀번호 중 무엇과 일치하는지 확인.

    Returns
    -------
    str | None
        "admin" 또는 "staff" — 로그인 성공 시 역할 문자열.
        아직 로그인 전이면 None (호출하는 쪽에서 st.stop()으로 이어감).
    """
    if st.session_state.get("app_role"):
        return st.session_state.app_role

    login_logger = get_login_sheet_logger()

    if login_logger.enabled:
        # 구글시트 기반 — "최초 계정 설정"이 아직 안 끝났으면 그 화면을 보여줌
        if not login_logger.is_account_setup_done():
            render_first_time_account_setup(login_logger)
            return None

        login_wrap = st.container(key="pf_login_wrap")
        with login_wrap:
            st.markdown(
                '<div class="pf-login-panel">'
                '<p class="pf-login-eyebrow">사내 로그인</p>'
                '<p class="pf-login-heading">비밀번호를 입력해주세요</p>'
                '<p class="pf-login-desc">포스원 회계법인 직원 전용 화면입니다.<br>'
                '관리자(회계사) 또는 직원 비밀번호를 입력해주세요.</p>'
                '</div>',
                unsafe_allow_html=True,
            )

            with st.form("app_login_form"):
                pw_input = st.text_input("비밀번호", type="password", label_visibility="collapsed")
                submitted = st.form_submit_button("입장", type="primary")

        if submitted:
            if login_logger.verify_account_password("admin", pw_input):
                st.session_state.app_role = "admin"
                st.rerun()
            elif login_logger.verify_account_password("staff", pw_input):
                st.session_state.app_role = "staff"
                st.rerun()
            elif pw_input.strip() and pw_input.strip() == _get_master_reset_password():
                # 비상 복구 비밀번호 (2026-07-09 추가):
                # 계정설정 탭의 정상 비밀번호를 잊어버렸거나(원인 불명 포함)
                # 잠긴 경우를 대비한 "항상 통하는" 백업 열쇠. 기본값은 "4119"이며
                # .env/Secrets의 MASTER_RESET_PASSWORD로 바꿀 수 있음. 이 값으로
                # 들어오면 관리자 권한을 주되, 화면에 눈에 띄게 경고를 남겨 즉시
                # 정식 비밀번호로 바꾸도록 유도함(그대로 두면 보안 구멍이므로).
                st.session_state.app_role = "admin"
                st.session_state["_logged_in_via_master_reset"] = True
                st.rerun()
            else:
                st.error("비밀번호가 일치하지 않습니다.")

        return None

    # 구글시트 로깅이 비활성 상태인 경우: 하위 호환을 위해 환경변수 방식으로 폴백
    admin_password = os.getenv("ADMIN_PASSWORD", "").strip() or os.getenv("APP_PASSWORD", "").strip()
    staff_password = os.getenv("STAFF_PASSWORD", "").strip()

    if not admin_password:
        st.warning(
            "구글시트 연동이 비활성 상태이고, ADMIN_PASSWORD(또는 기존 APP_PASSWORD)도 "
            "설정되어 있지 않습니다. 로컬 환경에서는 관리자 권한으로 그대로 진행되지만, "
            "웹에 배포할 때는 구글시트 연동을 설정하거나 .env/Secrets에 ADMIN_PASSWORD와 "
            "STAFF_PASSWORD를 반드시 설정해야 합니다."
        )
        st.session_state.app_role = "admin"
        return "admin"

    login_wrap = st.container(key="pf_login_wrap")
    with login_wrap:
        st.markdown(
            '<div class="pf-login-panel">'
            '<p class="pf-login-eyebrow">사내 로그인</p>'
            '<p class="pf-login-heading">비밀번호를 입력해주세요</p>'
            '<p class="pf-login-desc">포스원 회계법인 직원 전용 화면입니다.<br>'
            '관리자(회계사) 또는 직원 비밀번호를 입력해주세요.</p>'
            '</div>',
            unsafe_allow_html=True,
        )

        with st.form("app_login_form"):
            pw_input = st.text_input("비밀번호", type="password", label_visibility="collapsed")
            submitted = st.form_submit_button("입장", type="primary")

    if submitted:
        if pw_input == admin_password:
            st.session_state.app_role = "admin"
            st.rerun()
        elif staff_password and pw_input == staff_password:
            st.session_state.app_role = "staff"
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")

    return None


app_role = check_app_password()
if not app_role:
    st.stop()  # 비밀번호가 맞을 때까지 아래 모든 코드(엔진 초기화, 화면 등)를 실행하지 않음

# 화면 곳곳에서 "지금 로그인한 사람이 관리자인지"를 간단히 체크하기 위한 변수.
# 검색 기록 로깅 시 "사용자구분" 컬럼에도 이 값을 그대로 사용함(아래 사이드바
# 섹션의 user_type 자리를 대체).
is_admin = app_role == "admin"
user_type = "회계사" if is_admin else "직원"

if st.session_state.get("_logged_in_via_master_reset"):
    st.error(
        "⚠️ 비상 복구용 기본 비밀번호로 로그인했습니다. 보안을 위해 지금 바로 "
        "사이드바의 '로그인 비밀번호 관리'에서 관리자/직원 비밀번호를 새로 설정해주세요."
    )


# ----------------------------------------------------------------------
# 엔진 초기화 (세션당 1회, 캐시)
# ----------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_engine():
    return TaxAdvisorEngine()


try:
    engine = get_engine()
except (FileNotFoundError, ValueError) as e:
    st.error(str(e))
    st.stop()


# ----------------------------------------------------------------------
# 세션 상태
# ----------------------------------------------------------------------
# current_thread: 지금 화면에 보이는 "진행 중인 대화 묶음" (꼬리질문이 쌓이는 곳)
#   list of {"question": str, "answer": str, "time": str}
# current_thread: 지금 화면에 보이는 "진행 중인 대화 묶음" (꼬리질문이 쌓이는 곳)
#   list of {"question": str, "answer": str, "time": str}
st.session_state.setdefault("current_thread", [])

# backlog: "새 주제 시작"으로 넘어간 과거 묶음들
#   list of {"started_at": str, "turns": [...]}
st.session_state.setdefault("backlog", [])

st.session_state.setdefault("summary_doc", None)
st.session_state.setdefault("summary_doc_turns_count", 0)
st.session_state.setdefault("summary_doc_label", "")
st.session_state.setdefault("summary_doc_source_ts", None)
st.session_state.setdefault("kb_confirm_target", None)
# "저장 옵션"(사이드바) 위젯들의 기본값. 이 위젯들이 실제로 사이드바에서
# 그려지기 전에(예: st.dialog 팝업 안에서 저장 버튼을 먼저 누르는 경우)
# offer_downloads()가 이 값을 조회해도 NameError 없이 기본값을 쓰도록 함.
st.session_state.setdefault("save_formats", ["md"])
st.session_state.setdefault("use_custom_filename", False)
st.session_state.setdefault("custom_filename", "")
# "확정 저장 실행" 성공 직후 st.rerun()으로 화면을 정리할 때, 성공 메시지가
# rerun과 함께 사라지지 않도록 잠깐 담아두는 값 (아래에서 한 번 보여주고 비움)
st.session_state.setdefault("_kb_last_saved_path", None)


def render_copy_button(text: str, key: str, label: str = "복사"):
    """
    주어진 텍스트를 클립보드에 복사하는, 눈에 크게 띄지 않는 작은 아이콘 버튼을 그림.

    Streamlit의 st.button은 서버(파이썬) 쪽에서 눌림 여부만 알 수 있을 뿐,
    브라우저 클립보드에 직접 접근할 방법이 없음. 클립보드 복사는 브라우저
    JS가 있어야만 가능하므로, 순수 HTML/JS를 iframe에 삽입해 처리함(서버
    왕복 없이 클릭 즉시 브라우저에서 바로 실행됨).

    렌더링 방식 (2026-07-02 변경): 기존에는 st.components.v1.html을 썼는데,
    Streamlit이 이 API를 2026-06-01 이후 제거 예정으로 공지하고 있어
    (배포 로그에 "Please replace st.components.v1.html with st.iframe"
    경고가 뜸), st.iframe으로 교체함. st.iframe은 HTML 문자열을 넣으면
    자동으로 감지해 같은 방식으로 렌더링해줌(사용법 동일).

    스타일 (2026-07-02 변경): 처음엔 네이비/골드 배경의 큰 버튼이었는데,
    실사용 피드백으로 "너무 눈에 띈다"는 의견을 받아 — 지금 앱에서 이미
    "다음 라운드에 보낼 질문"을 보여줄 때 쓰는 st.code() 코드블록 우측
    상단의 작은 복사 아이콘과 비슷한 느낌으로, 테두리만 있는 작은
    아이콘 버튼으로 축소함.

    복사 방식은 두 단계로 시도함:
    ① navigator.clipboard.writeText — 최신 API. 다만 iframe 안에서
       실행되다 보니, 배포 환경에 따라 브라우저의 Permissions Policy가
       iframe에는 클립보드 쓰기 권한을 기본적으로 내려주지 않아 조용히
       실패하는 사례가 확인됨(실사용 테스트 — 버튼은 보이지만 클릭해도
       복사가 안 되는 현상).
    ② ①이 실패하면 즉시 document.execCommand('copy')로 대체 — 화면
       밖에 숨겨진 textarea에 텍스트를 넣고 선택한 뒤 복사하는 오래된
       방식이라 최신 API보다 지원 범위가 넓고, iframe 안에서도 별도
       권한 없이 대부분 동작함.

    text는 json.dumps로 감싸 JS 문자열 리터럴로 안전하게 이스케이프함
    (텍스트 안에 따옴표·줄바꿈·백틱 등이 있어도 깨지지 않도록).

    key는 버튼의 HTML id로 쓰이므로, 화면에 같은 컴포넌트가 여러 개 있을
    때는 반드시 서로 다른 값을 넘겨야 함(다른 위젯들과 동일한 패턴).
    """
    safe_text = json.dumps(text or "")
    safe_key = re.sub(r"[^0-9a-zA-Z_]", "_", str(key))
    btn_id = f"pf_copy_btn_{safe_key}"
    ta_id = f"pf_copy_ta_{safe_key}"
    html = f"""
    <div style="display:flex; justify-content:flex-end; margin: 0 0 4px;">
      <textarea id="{ta_id}" style="position:fixed; top:-9999px; left:-9999px; opacity:0;"></textarea>
      <button id="{btn_id}" onclick="pfCopy_{safe_key}()" title="복사" style="
        background-color: transparent;
        color: #6b7280;
        border: 1px solid #d1d5db;
        border-radius: 5px;
        padding: 1px 8px;
        font-size: 0.72rem;
        font-weight: 500;
        cursor: pointer;
        font-family: inherit;
        line-height: 1.6;
      ">📋 {label}</button>
    </div>
    <script>
      function pfCopy_{safe_key}() {{
        var text = {safe_text};
        var btn = document.getElementById('{btn_id}');
        var originalHTML = btn.innerHTML;

        function onSuccess() {{
          btn.innerHTML = '✅';
          setTimeout(function() {{ btn.innerHTML = originalHTML; }}, 1200);
        }}
        function onFail() {{
          btn.innerHTML = '복사 실패';
          setTimeout(function() {{ btn.innerHTML = originalHTML; }}, 1800);
        }}
        function fallbackCopy() {{
          try {{
            var ta = document.getElementById('{ta_id}');
            ta.value = text;
            ta.style.top = '0px';
            ta.focus();
            ta.select();
            var ok = document.execCommand('copy');
            ta.style.top = '-9999px';
            if (ok) {{ onSuccess(); }} else {{ onFail(); }}
          }} catch (e) {{
            onFail();
          }}
        }}

        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(text).then(onSuccess).catch(fallbackCopy);
        }} else {{
          fallbackCopy();
        }}
      }}
    </script>
    """
    st.iframe(html, height=30)


def render_copyable_text(text: str, key: str, label: str = "복사"):
    """
    텍스트를 마크다운으로 표시하고, 그 위 우측에 작은 복사 아이콘 버튼을 함께 그림.
    (드래그 + Ctrl+C 대신 이 버튼을 쓰면, Streamlit의 'c' 단축키가
    실수로 눌려 캐시가 지워지고 세션이 끊기는 문제를 피할 수 있음)
    """
    render_copy_button(text, key=key, label=label)
    st.markdown(text)


def get_thread_label(turns):
    """묶음의 첫 질문을 짧게 잘라 제목처럼 사용"""
    if not turns:
        return "(빈 대화)"
    first_q = turns[0]["question"]
    return first_q[:40] + ("..." if len(first_q) > 40 else "")


def offer_downloads(title: str, question: str, answer: str, filename_hint: str = None, key_prefix: str = ""):
    """
    결과를 md/docx/pdf 바이트로 만들어 다운로드 버튼들을 그려줌.

    설계 의도 (2026-06-25 — 웹 배포 지원):
    - 기존 do_save()는 서버의 로컬 폴더(output_dir)에 직접 파일을 써서 저장했음.
      PC에서는 동작했지만, 웹 서버에는 그 경로가 없어서 폴더 생성 자체가
      실패하고, 그 예외가 처리되지 않은 채 위로 던져지면서 화면 전체가
      다시 그려져 "펼쳐둔 화면이 갑자기 초기화면처럼 닫혀버리는" 문제가 있었음.
    - 해결: 디스크에 쓰지 않고 바이트만 만들어서, st.download_button으로
      사용자 컴퓨터에 직접 내려주는 방식으로 전환함. 이러면 서버 경로 문제가
      원천적으로 없어지고, 로컬 PC에서 쓰든 웹에서 쓰든 동일하게 동작함.
    - 다운로드는 버튼을 누르는 사용자의 즉시 행동이므로, 별도의 "저장 폴더"
      설정이 필요 없음. 어디에 받을지는 브라우저의 다운로드 동작(또는
      다운로드 위치 선택 대화상자)에 맡김.

    Parameters
    ----------
    key_prefix : str
        같은 화면에 여러 개의 다운로드 버튼이 있을 때 Streamlit 위젯 키가
        충돌하지 않도록 구분하는 접두사 (예: 호출 위치 + 타임스탬프 조합).
    """
    # 버그 수정 메모 (2026-07-02, 2차): 이 함수 자체가 사이드바 코드보다
    # 뒤쪽에 정의되어 있었음. st.dialog(검색 기록 상세보기 등)는 열리는
    # 순간 그 아래 스크립트 실행을 멈추는 특성이 있어서, 사이드바보다 먼저
    # 다이얼로그가 뜨면 "def offer_downloads(...)" 라인 자체가 이번 실행에서
    # 아직 한 번도 지나가지 않은 상태가 되어 "NameError: name
    # 'offer_downloads' is not defined"가 발생했음(사이드바 위젯 변수들
    # (save_formats 등)에서 겪었던 것과 같은 종류의 문제가 함수 자체에도
    # 있었던 것). 함수 정의를 사이드바보다 훨씬 앞쪽(다른 헬퍼 함수들
    # 근처)으로 옮겨, 어떤 다이얼로그가 언제 열리든 항상 정의되어 있는
    # 상태가 되도록 함.
    #
    # 안의 save_formats 등 값 자체는 여전히 session_state에서 읽음(1차
    # 수정 때와 동일) — 실행 순서와 무관하게 항상 안전하도록 이중으로 보강.
    use_custom_filename = st.session_state.get("use_custom_filename", False)
    custom_filename = st.session_state.get("custom_filename", "")
    filename = safe_filename(
        custom_filename.strip() if (use_custom_filename and custom_filename.strip()) else filename_hint
    )
    formats = st.session_state.get("save_formats") or ["md"]

    built = build_export_bytes(title=title, question=question, response_md=answer, formats=formats)

    ext_map = {"md": ("text/markdown", ".md"), "docx": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"), "pdf": ("application/pdf", ".pdf")}
    cols = st.columns(len(formats)) if formats else []
    for col, fmt in zip(cols, formats):
        data = built.get(fmt)
        with col:
            if data:
                mime, ext = ext_map[fmt]
                st.download_button(
                    label=f"{fmt.upper()} 다운로드",
                    data=data,
                    file_name=f"{filename}{ext}",
                    mime=mime,
                    key=f"{key_prefix}_dl_{fmt}",
                )
            else:
                st.button(f"{fmt.upper()} 다운로드 (사용 불가)", disabled=True, key=f"{key_prefix}_dl_{fmt}_disabled")

    if built["errors"]:
        for err in built["errors"]:
            st.caption(f"⚠ {err}")


def render_confirm_to_kb_button(
    question: str,
    answer: str,
    key_prefix: str,
    dialog_row_key: str = None,
    source_summary_timestamp: str = None,
    source_log_timestamp: str = None,
    already_confirmed: bool = False,
):
    """
    '지식베이스에 확정 저장' 버튼만 그림. 누르면 실제 작업 화면은 이 버튼이 있는
    좁은 위치(다이얼로그, 카드 등)가 아니라 메인 화면 하단의 전용 영역에서 펼쳐짐
    (render_confirm_to_kb_workspace 참고).

    설계 의도 (2026-06-25 추가 — 작업 공간 분리):
    - 처음에는 이 흐름 전체(검증 결과 + 큰 텍스트 수정칸 + 파일선택 + PIN)를 버튼이
      있는 그 자리(다이얼로그 안, 좁은 카드 안)에 그대로 펼쳤었음. 하지만 답변
      내용이 길 때(200줄 가까이) 좁은 영역의 작은 텍스트칸으로는 수정하기 너무
      불편하다는 피드백을 받음.
    - 해결: 버튼은 그 자리에 작게 두고, 누르면 "이 질문/답변을 지금부터 확정 저장
      작업 대상으로 지정"만 함. 실제 검증 결과 표시, 수정용 텍스트칸, 파일 선택,
      PIN 입력은 메인 화면(다이얼로그 바깥) 맨 아래의 별도의 큰 섹션에서 진행됨.
      그 섹션은 화면 전체 너비를 쓸 수 있어 긴 내용도 편하게 수정 가능함.

    버그 수정 (2026-06-25 추가 — 다이얼로그 안에서 호출 시):
    - @st.dialog로 띄운 다이얼로그 안에서 이 버튼을 누르면, session_state만
      바꾸는 것으로는 다이얼로그가 닫히지 않아 메인 화면 맨 아래의 작업 공간이
      나타나지 않는 문제가 있었음(다이얼로그가 열려있는 동안 다이얼로그 바깥
      영역이 갱신되지 않는 Streamlit의 동작 특성 때문). 또한 이 상태에서 작업을
      계속 진행하면 화면이 초기화면으로 튕기는 현상도 발생했음.
    - 해결: dialog_row_key(예: "_dialog_log_row")를 전달받으면, 그 값을 None으로
      비워서 다음 리런부터 다이얼로그가 다시 열리지 않게 하고, st.rerun()으로
      명시적으로 다이얼로그를 닫음. 이후 메인 화면이 정상적으로 다시 그려지면서
      작업 공간이 나타남.

    Parameters
    ----------
    source_summary_timestamp : str, optional
        이 확정 대상이 '종합문서' 탭의 특정 행에서 온 것이면 그 행의 '일시' 값을
        전달함(2026-07-09 추가). 확정이 완료되면 지식베이스에 새 항목을 추가하는
        것과 별개로, 이 값을 이용해 원본 종합문서 행 자체도 확정된 최종본으로
        덮어씀(update_summary) — 검증 전 원본과 확정된 최종본이 따로 남아
        헷갈리는 것을 막기 위함. 개별 질문/답변 확정 시에는 전달하지 않음(그
        경우는 검색기록과 지식베이스가 계속 별도로 남는 기존 방식 그대로 유지).
    source_log_timestamp : str, optional
        이 확정 대상이 개별 질문/답변(검색기록 탭의 특정 행)이면 그 행의
        '일시' 값을 전달함(2026-07-09 추가). 종합문서와 달리 내용을 덮어쓰지는
        않고, 검색기록 탭의 '확정여부' 컬럼만 "확정됨"으로 표시함
        (mark_log_confirmed) — 검색기록에서 이미 지식베이스에 올라간
        질문인지 한눈에 알 수 있도록 하기 위함.
    """
    target_key = "kb_confirm_target"
    # 이미 지식베이스에 확정된 문서라면 '확정 저장' 버튼을 다시 보여주지 않는다.
    # (이미 확정 배지가 떠 있는데 아래에 또 확정 버튼이 있으면 "이미 확정한 걸
    #  또 확정"하게 되어 혼란스럽다는 피드백 반영.) 수정이 필요하면 새 질의로
    # 다시 확정하는 흐름을 쓰도록 안내만 남긴다.
    if already_confirmed:
        st.caption(
            "이미 지식베이스에 확정된 문서입니다. (내용을 바꿔야 하면 새로 질의해 "
            "다시 확정하세요.)"
        )
        return
    if st.button("지식베이스에 확정 저장", key=f"{key_prefix}_confirm_entry_btn"):
        st.session_state[target_key] = {
            "question": question,
            "answer": answer,
            "key_prefix": key_prefix,
            # v1.6 추가 — 교차검증대기 구글시트 백업용 세션ID. 같은 작업에서
            # 나온 모든 라운드(1차,2차...)가 이 ID로 묶여 저장/복원/삭제됨.
            "crosscheck_session_id": str(uuid.uuid4()),
            "source_summary_timestamp": source_summary_timestamp,
            "source_log_timestamp": source_log_timestamp,
        }
        # 이전에 다른 항목을 검증/수정하던 상태가 남아있으면 깨끗하게 초기화.
        # edited_content 키 자체를 세션 상태에서 제거함(아직 확정 전이므로
        # 값이 존재하지 않아야 정상 — 그래야 확정 단계에서 "키가 없다"는
        # 조건이 True가 되어 최신 라운드 문서가 정상적으로 채워짐).
        st.session_state.pop(f"{key_prefix}_edited_content", None)

        # v1.7 — 라운드 진행 관련 상태도 새 대상 지정 시 깨끗하게 초기화.
        # 신규 질문이든 검색기록에서 불러온 과거 질문이든, 이 버튼을 누르는
        # 순간부터는 항상 "처음부터 새로 시작하는 라운드 진행"이어야 함 —
        # 이전에 다른 질문을 검증하다 만든 기록이 섞여 들어가면 안 됨.
        st.session_state.pop(f"{key_prefix}_current_doc", None)
        st.session_state.pop(f"{key_prefix}_doc_rounds", None)
        st.session_state.pop(f"{key_prefix}_next_prompt", None)
        st.session_state.pop(f"{key_prefix}_proceed_to_confirm", None)
        st.session_state.pop(f"{key_prefix}_recommended_file", None)
        st.session_state.pop(f"{key_prefix}_recommended_reason", None)
        # v1.11 — Claude 자동 교차검증 관련 상태도 함께 초기화
        st.session_state.pop(f"{key_prefix}_auto_msg", None)
        st.session_state.pop(f"{key_prefix}_auto_final_critique", None)

        if dialog_row_key:
            # 다이얼로그 안에서 호출된 경우: 다이얼로그를 닫고 메인 화면으로 이동
            st.session_state[dialog_row_key] = None
            st.rerun()
        else:
            # 버그 수정 (2026-06-27 — "1차 답변 화면에서만 튕기는" 문제):
            # 이 분기는 그동안 st.rerun()을 호출하지 않고 st.info()만 출력했음.
            # st.button()이 True가 되는 순간 Streamlit이 이미 "이 실행이 끝나면
            # 다시 그려라"를 예약해두는데, 그 사이 위에서 session_state를 여러 개
            # 바꿔놓은 상태로(kb_confirm_target 설정, edited_content 제거 등)
            # 이번 실행의 나머지 부분(같은 반복문의 다른 항목들, 그 아래 "전체
            # 종합" 섹션 등)이 그대로 이어서 실행됨. 이 어중간한 상태로 나머지
            # 스크립트가 끝까지 실행되다 예외가 나면 페이지가 초기 화면으로
            # 튕기는 현상으로 이어질 수 있음(다이얼로그 경로는 st.rerun()을
            # 명시적으로 호출해 이런 문제가 없었음 — 두 경로의 유일한 차이).
            # 해결: 다이얼로그 경로와 동일하게 st.rerun()을 명시적으로 호출해,
            # session_state 변경 직후 깨끗하게 다시 시작하도록 통일함.
            st.info("화면 맨 아래 '지식베이스 확정 저장 작업 공간'으로 이동해 진행해주세요.")
            st.rerun()


def render_confirm_to_kb_workspace():
    """
    실제 검증 → 수정 → 저장 흐름이 펼쳐지는 전용 작업 공간.
    render_confirm_to_kb_button으로 대상이 지정된 경우에만 화면에 나타나며,
    메인 화면 맨 아래(화면 전체 너비)에 한 번만 그려짐.

    구조 변경 (v1.7 — 2026-06-28, "본문이 매 라운드 보이고 바뀌는" 구조로 전환):
    - v1.6의 교차검증 스레드는 라운드를 여러 번 돌려도 "검증 텍스트"만
      누적되고, 본문(실제 답변 문서)은 맨 마지막 "확정 단계로 진행"을 누를
      때 딱 한 번만 바뀌는 구조였음. 회계사 피드백: "라운드를 여러 번
      거쳤는데 실제 결과물은 1차 검증분만 반영되어 부실하다", "본문이 안
      보이니 뭐가 바뀌고 있는지 예측이 안 된다".
    - 변경: 1라운드는 기존처럼 "① 웹검색으로 검증하기" 버튼으로 웹검색
      동반 검증+자동수정을 한 번에 수행해 "확정 문서 v1"을 만듦. 이후
      라운드부터는 웹검색 없이(회계사 피드백 반영 — 웹검색은 느려서 1회만),
      외부 AI의 자유 형식 의견을 받아 문서 전체를 다시 쓰는 방식
      (apply_external_feedback)으로 v2, v3...을 만듦.
    - 매 라운드 "그 시점의 확정 문서 전체"와 "이번에 무엇을 고쳤는지"가
      화면에 그대로 보이고, 다음 라운드에 외부 AI에게 보낼 질문에도 현재
      문서 전체가 포함됨(검증 텍스트 조각이 아니라 완성된 문서를 보고
      교차검증하게 하여, 문서 전체에 걸친 일관성 문제도 잡아내게 함).
    - 신규 질문이든 검색기록(구글시트)에서 불러온 과거 질문이든, 이 흐름은
      항상 동일하게 동작함.
    - 회계사가 만족하는 시점에 "이걸로 확정" 버튼을 누르면, 그 시점의 문서
      버전을 그대로 가지고 곧바로 ③ 파일선택/PIN 단계로 넘어감 — 별도의
      "자동 수정" 단계가 더 필요하지 않음(이미 매 라운드 본문에 반영됨).
    """
    target = st.session_state.get("kb_confirm_target")
    if not target:
        return

    question = target["question"]
    answer = target["answer"]
    key_prefix = target["key_prefix"]
    # 교차검증대기 백업용 세션ID. render_confirm_to_kb_button에서 만들어
    # 넣어두지만, 혹시 그 경로를 거치지 않고 target이 만들어진 경우(예:
    # 과거 버전과의 호환)를 대비해 없으면 여기서 즉석 생성함.
    crosscheck_session_id = target.get("crosscheck_session_id")
    if not crosscheck_session_id:
        crosscheck_session_id = str(uuid.uuid4())
        target["crosscheck_session_id"] = crosscheck_session_id
    # 종합문서 확정 시 원본 '종합문서' 탭 행을 덮어쓰기 위한 타임스탬프
    # (개별 질문/답변 확정 시에는 None — render_confirm_to_kb_button 참고)
    source_summary_timestamp = target.get("source_summary_timestamp")
    source_log_timestamp = target.get("source_log_timestamp")
    st.session_state["kb_confirm_target"] = target

    st.divider()
    st.header("지식베이스 확정 저장 작업 공간")
    st.caption("아래에서 검증 결과를 확인하고, 필요한 부분을 직접 수정한 뒤 PIN으로 최종 승인하세요.")

    with st.expander("대상 질문 / 원본 답변 보기", expanded=False):
        st.markdown(f"**질문**: {question}")
        st.markdown("**원본 답변**")
        render_copyable_text(answer, key=f"{key_prefix}_orig_answer")

    if not engine.has_pin_set():
        st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
        return

    edited_key = f"{key_prefix}_edited_content"
    # v1.7 라운드 상태 키들
    rounds_key = f"{key_prefix}_doc_rounds"        # list[dict] — 완료된 라운드 기록
    current_doc_key = f"{key_prefix}_current_doc"  # str — 지금 시점의 확정 문서 전체
    next_prompt_key = f"{key_prefix}_next_prompt"  # str — 다음 라운드에 보낼 질문
    proceed_key = f"{key_prefix}_proceed_to_confirm"  # bool — "이걸로 확정" 눌렀는지

    # ------------------------------------------------------------------
    # 1라운드: 원본 답변을 그대로 "교차검증 대상 문서 v1"으로 준비
    # ------------------------------------------------------------------
    # v1.11.2 변경(B안): 기존에는 여기서 Gemini가 웹검색으로 자체검증하고
    # 본문을 자동수정(verify_before_confirm)해 1차 문서를 만들었으나, 이
    # 자체검증 단계가 오히려 "맞는 값을 틀리게" 바꾸는 사례가 확인됨(예:
    # 납부지연가산세율 2.2/10,000을 2.0으로 바꾸고 없는 시행일을 생성).
    # 저자 모델(Gemini)이 자기 답을 스스로 검증하면 같은 사각지대에 갇힐 뿐
    # 아니라 새 오류를 만들 수 있어, 이 자체검증을 제거하고 원본 답변을 곧바로
    # 독립 검증자(Claude) 교차검증에 넘긴다. 검증 후 본문 수정(반영)은 이후에도
    # 기존처럼 Gemini(apply_external_feedback)가 담당한다.
    if st.session_state.get(current_doc_key) is None:
        if st.button("① 교차검증 시작 (원본 답변을 검증 대상으로 준비)", key=f"{key_prefix}_run_verify_btn", type="primary"):
            # 저장 파일 추천만 가볍게 받아둔다(문서 내용은 건드리지 않음, 웹검색 없음).
            with st.spinner("교차검증 대상 문서를 준비하는 중입니다..."):
                try:
                    rec_file, rec_reason = engine._recommend_knowledge_file(question, answer)
                except Exception:
                    rec_file, rec_reason = "", ""

            _init_summary = "- 원본 답변 (Gemini 자체검증 없이 교차검증 대상으로 사용)"
            st.session_state[current_doc_key] = answer
            st.session_state[rounds_key] = [{
                "round": 1,
                "document": answer,
                "change_summary": _init_summary,
                "verification_detail": "",
                "next_prompt": "",  # 아래에서 채움
                "external_ai_input": "",  # 다음 라운드에서 답변 오면 채움
            }]
            # 추천 파일/이유(저장 대상 파일 선택용) — 문서 내용과 무관한 분류 결과
            st.session_state[f"{key_prefix}_recommended_file"] = rec_file
            st.session_state[f"{key_prefix}_recommended_reason"] = rec_reason

            next_prompt = engine.build_next_round_prompt(question, answer, "")
            st.session_state[next_prompt_key] = next_prompt
            st.session_state[rounds_key][0]["next_prompt"] = next_prompt

            # 안전장치 — WebSocket 끊김 대비, 1라운드 준비 상태를 곧바로 백업
            if engine.sheet_logger and engine.sheet_logger.enabled:
                engine.sheet_logger.save_crosscheck_round(
                    session_id=crosscheck_session_id,
                    round_no=1,
                    question=question,
                    original_answer=answer,
                    verification_text=_init_summary,
                    cross_prompt=next_prompt,
                    external_ai_input="",
                    next_verification_text=answer,
                )
            st.rerun()
        else:
            st.caption(
                "원본 답변을 그대로 교차검증 대상으로 준비합니다(Gemini 자체검증은 "
                "하지 않습니다). 준비 후 아래에서 Claude 자동 교차검증을 돌리거나, "
                "다른 AI에게 수동으로 교차검증을 요청해 문서를 다듬을 수 있습니다. "
                "충분하다면 곧바로 확정할 수도 있습니다."
            )
            return

    doc_rounds = st.session_state.get(rounds_key, [])
    current_document = st.session_state[current_doc_key]
    current_round = len(doc_rounds)  # 지금까지 완료된 라운드 수 (= 가장 최근 라운드 번호)

    st.markdown("### ② 라운드별 진행 내역")
    st.caption(f"총 {current_round}차까지 진행되었습니다. 라운드마다 그 시점의 확정 문서 전체가 아래에 누적되어 표시됩니다.")

    # 완료된 라운드를 위에서부터 차례로 표시 (가장 최근 라운드만 문서를 펼쳐서 보여줌)
    for r in doc_rounds:
        st.divider()
        st.markdown(f"#### 🔹 {r['round']}차 확정 문서")
        st.markdown(f"**이번 라운드에 반영된 수정 사항**")
        st.write(r["change_summary"])
        is_latest = (r["round"] == current_round)
        with st.expander(f"{r['round']}차 확정 문서 전체 보기", expanded=is_latest):
            render_copyable_text(r["document"], key=f"{key_prefix}_round_{r['round']}_doc")
        if r.get("external_ai_input"):
            with st.expander(f"{r['round']}차에 보냈던 질문", expanded=False):
                st.code(r.get("next_prompt", "(기록 없음)"), language="text")
            _src = r.get("source", "")
            if _src == "auto":
                st.markdown(f"**{r['round']}차 — 🤖 Claude 자동 교차검증이 지적한 내용**")
            elif _src == "manual":
                st.markdown(f"**{r['round']}차 — ✍️ 수동으로 붙여넣은 다른 AI의 답변**")
            else:
                st.markdown(f"**{r['round']}차 검증 후 받은 답변**")
            st.success(r["external_ai_input"])

    if not st.session_state.get(proceed_key):
        st.divider()
        tab_auto, tab_manual = st.tabs(["🤖 자동 (Claude)", "✍️ 수동 (다른 AI 복사/붙여넣기)"])

        # ══════════════════════════════════════════════════════════════
        # 탭 1 — Claude 자동 교차검증
        # ══════════════════════════════════════════════════════════════
        # 서로 다른 벤더인 Claude(Sonnet)를 API로 직접 호출해 "점검 → 본문
        # 반영"을 자동으로 여러 라운드 반복함. 종료조건은 ①Claude '완료' 신호
        # ②본문 무변화(수렴) ③최대 라운드 3중(엔진 참고).
        with tab_auto:
            # 직전 자동 실행 결과 메시지 / Claude 최종 판단 표시
            _auto_msg = st.session_state.get(f"{key_prefix}_auto_msg")
            if _auto_msg:
                _lvl, _txt = _auto_msg
                getattr(st, _lvl, st.info)(_txt)
                _fc = st.session_state.get(f"{key_prefix}_auto_final_critique", "")
                if _fc:
                    # 성공(✅)이 아니면(경고·실패) Claude의 남은 지적을 놓치지 않도록
                    # 최종 의견을 자동으로 펼쳐서 보여준다.
                    _expand_critique = (_lvl != "success")
                    with st.expander("Claude의 마지막 검토 의견 전체 보기", expanded=_expand_critique):
                        render_copyable_text(_fc, key=f"{key_prefix}_auto_final_critique_view")

            if not engine.auto_cross_check_enabled:
                st.info(
                    "자동 교차검증(Claude)을 쓰려면 Streamlit Secrets에 ANTHROPIC_API_KEY를 "
                    "추가하세요(공개 저장소이므로 코드에는 넣지 마세요). 지금은 '✍️ 수동' "
                    "탭으로 진행할 수 있습니다."
                )
            else:
                st.caption(
                    "서로 다른 벤더인 Claude(Sonnet)가 조문을 우선 근거로 문서를 점검하고 "
                    "그 지적을 문서에 자동 반영합니다. ① Claude가 '더 고칠 것 없음'으로 "
                    "판단하거나 ② 문서가 더 이상 바뀌지 않거나 ③ 최대 라운드에 도달하면 "
                    "자동으로 멈춥니다."
                )
                _col_r, _col_w = st.columns([1, 1])
                with _col_r:
                    _max_rounds = st.number_input(
                        "최대 라운드 수",
                        min_value=1, max_value=6, value=2, step=1,
                        key=f"{key_prefix}_auto_max_rounds",
                        help="Claude 호출 상한입니다. 보통 1~2회면 핵심 오류가 잡힙니다. "
                             "라운드가 늘수록 비용도 비례해서 늘어납니다.",
                    )
                with _col_w:
                    _use_web = st.checkbox(
                        "웹검색 사용",
                        value=False,
                        key=f"{key_prefix}_auto_use_web",
                        help="켜면 Claude가 조문으로 부족할 때 웹을 검색합니다. 정확도가 "
                             "조금 오를 수 있으나 호출당 별도 과금 + 토큰이 늘어 비용이 "
                             "크게 증가합니다. 조항 번호·특례 등 핵심 오류는 웹검색 없이도 "
                             "대부분 잡히니 기본은 꺼두는 것을 권합니다.",
                    )
                st.caption(
                    "⚠️ API 호출은 응답이 비정상 종료돼도 과금될 수 있습니다. 실행 중 "
                    "오류가 나면 콘솔 Billing에서 사용량을 한 번 확인해보세요."
                )
                if st.button(
                    "🤖 Claude로 자동 교차검증 실행",
                    key=f"{key_prefix}_auto_run_btn_{current_round}",
                    type="primary",
                ):
                    with st.spinner(
                        "Claude가 문서를 검토하고 지적을 반영하는 과정을 반복하는 중입니다... "
                        "(라운드당 최대 수십 초, 웹검색 시 더 걸릴 수 있음)"
                    ):
                        _counter = {"n": current_round}

                        def _backup_step(step):
                            # 매 라운드 결과를 곧바로 구글시트에 백업(중간에 연결이
                            # 끊겨도 진행분이 남도록). 실패해도 루프는 계속.
                            _counter["n"] += 1
                            if engine.sheet_logger and engine.sheet_logger.enabled:
                                try:
                                    engine.sheet_logger.save_crosscheck_round(
                                        session_id=crosscheck_session_id,
                                        round_no=_counter["n"],
                                        question=question,
                                        original_answer=answer,
                                        verification_text=step["change_summary"],
                                        cross_prompt=step["prompt_sent"],
                                        external_ai_input=step["critique"],
                                        next_verification_text=step["new_document"],
                                    )
                                except Exception:
                                    pass

                        result = engine.run_auto_cross_check(
                            question=question,
                            current_document=current_document,
                            max_rounds=int(_max_rounds),
                            prev_change_summary=(doc_rounds[-1]["change_summary"] if doc_rounds else ""),
                            enable_web_search=bool(_use_web),
                            on_step=_backup_step,
                        )

                    # 반환된 각 라운드를 기존 수동 흐름과 동일한 방식으로 기록에 반영
                    for step in result["steps"]:
                        _rounds = st.session_state[rounds_key]
                        # 이 step을 보낸 시점의 마지막 라운드에 "보낸 질문/받은 답변"을 채움
                        _rounds[-1]["next_prompt"] = step["prompt_sent"]
                        _rounds[-1]["external_ai_input"] = step["critique"]
                        _rounds[-1]["source"] = "auto"
                        _rounds.append({
                            "round": len(_rounds) + 1,
                            "document": step["new_document"],
                            "change_summary": step["change_summary"],
                            "verification_detail": "",
                            "next_prompt": "",
                            "external_ai_input": "",
                        })
                        st.session_state[rounds_key] = _rounds
                        st.session_state[current_doc_key] = step["new_document"]

                    # 이후 수동 라운드를 이어갈 수 있도록 최종 문서 기준 질문 재생성
                    _rounds = st.session_state[rounds_key]
                    _final_doc = st.session_state[current_doc_key]
                    _last_summary = _rounds[-1]["change_summary"] if _rounds else ""
                    st.session_state[next_prompt_key] = engine.build_next_round_prompt(
                        question, _final_doc, _last_summary
                    )
                    if _rounds:
                        _rounds[-1]["next_prompt"] = st.session_state[next_prompt_key]

                    # 종료 결과 메시지 구성 (rerun 후 표시하기 위해 세션에 저장)
                    #
                    # 판정은 Claude의 6등급(S/A/B/C/D/F)을 기준으로 한다.
                    #   S/A = 확정 가능(초록) → 확정 권유
                    #   B/C = 확정 가능하나 회계사 판단(연두/노랑) → 확정 or 재실행 선택
                    #   D/F = 실질 오류(주황/빨강) → 계속 돌렸으나 상한/정체로 멈춘 상태
                    #   ?   = 등급 파싱 실패 → 회계사 직접 검토
                    _n = result["rounds_run"]
                    _applied = len(result["steps"])
                    _reason = result["stop_reason"]
                    _grade = result.get("final_grade", "")
                    _greason = (result.get("final_reason", "") or "").strip()
                    _gtail = f" — {_greason}" if _greason else ""

                    if _reason == "apply_failed":
                        _msg = ("error", f"❌ Claude가 지적한 내용을 문서에 자동 반영하지 못했습니다(반영 실패). 이 문서에는 Claude의 지적이 아직 반영되지 않았으니 그대로 확정하지 마세요. 아래 'Claude의 마지막 검토 의견'을 직접 확인해 수동으로 반영하거나 '✍️ 수동' 탭으로 진행하세요.")
                    elif _reason == "error":
                        if result["steps"]:
                            _msg = ("warning", f"⚠️ 자동 교차검증이 중간에 중단됐습니다 (중단 전 {_applied}회 반영): {result['error']}")
                        else:
                            _msg = ("error", f"❌ 자동 교차검증을 시작하지 못했습니다: {result['error']}")
                    elif _grade == "S":
                        _msg = ("success", f"✅ [S 확정] 더 볼 것이 없는 완성 상태입니다 (검토 {_n}회, 반영 {_applied}회). 확정하셔도 됩니다.{_gtail}")
                    elif _grade == "A":
                        _msg = ("success", f"✅ [A 확정가능] 확정해도 됩니다. 완벽하진 않으나 사실관계·결론·수치·조항이 모두 맞습니다 (검토 {_n}회, 반영 {_applied}회).{_gtail}")
                    elif _grade == "B":
                        _msg = ("success", f"🟢 [B 보통·개선실익 낮음] 확정하셔도 됩니다. 남은 사항은 사람이 원문으로 확인할 부분이라 더 돌려도 크게 나아지지 않습니다 (검토 {_n}회, 반영 {_applied}회). 확정하거나, 원하시면 한 번 더 돌릴 수 있습니다.{_gtail}")
                    elif _grade == "C":
                        _msg = ("warning", f"🟡 [C 보통·개선실익 있음] 확정해도 큰 문제는 없으나, 한 번 더 돌리면 더 정확해질 여지가 있습니다 (검토 {_n}회, 반영 {_applied}회). 확정할지 한 번 더 돌릴지 선택하세요.{_gtail}")
                    elif _grade in ("D", "F"):
                        # D/F인데 여기까지 왔다는 건 상한(max_rounds) 또는 정체(stalled)로 멈춘 것
                        _label = "F 중대오류" if _grade == "F" else "D 미흡"
                        if _reason == "stalled":
                            _msg = ("error", f"🟠 [{_label}] 실질 오류가 남아 있는데 수정이 더 진행되지 않고 정체됐습니다 (검토 {_n}회, 반영 {_applied}회). 그대로 확정하지 마시고, 아래 검토 의견을 보고 '✍️ 수동' 탭에서 직접 반영하세요.{_gtail}")
                        else:
                            _msg = ("error", f"🟠 [{_label}] 최대 {int(_max_rounds)}라운드에 도달했지만 실질 오류가 남아 있습니다 (반영 {_applied}회). 그대로 확정하지 마시고, 자동 실행을 한 번 더 누르거나 '✍️ 수동' 탭에서 직접 반영하세요.{_gtail}")
                    else:
                        # '?' 또는 예상 밖 — 회계사 직접 검토로 안전하게 넘김
                        _msg = ("warning", f"⚠️ Claude가 등급 판정을 형식대로 내리지 못했습니다 (검토 {_n}회, 반영 {_applied}회). 아래 'Claude의 마지막 검토 의견'을 직접 확인하고 판단하세요.{_gtail}")
                    st.session_state[f"{key_prefix}_auto_msg"] = _msg
                    st.session_state[f"{key_prefix}_auto_final_critique"] = result.get("final_critique", "")
                    st.rerun()

                st.caption("자동 반영이 끝난 뒤에도, '✍️ 수동' 탭에서 라운드를 더 이어가거나 자동 실행을 반복할 수 있습니다.")

        # ══════════════════════════════════════════════════════════════
        # 탭 2 — (기존) 수동 교차검증: 다른 AI에 복사해서 물어보고 답을 붙여넣음
        # ══════════════════════════════════════════════════════════════
        with tab_manual:
            st.markdown(f"#### 🔸 {current_round}차에 보낼 질문 (다음 라운드 준비)")
            st.caption(
                "아래 질문을 복사해서 다른 AI(ChatGPT, claude.ai 등)에게 물어보고, 받은 "
                "답변을 아래 칸에 붙여넣으면 그 의견을 반영해 문서를 다시 작성합니다(2차부터는 "
                "웹검색 없이 빠르게 처리됩니다). 이 과정은 원하는 만큼 반복할 수 있고, "
                "충분하다면 생략하고 바로 확정해도 됩니다. API 비용이 들지 않습니다."
            )
            if not st.session_state.get(next_prompt_key):
                st.session_state[next_prompt_key] = engine.build_next_round_prompt(
                    question, current_document, doc_rounds[-1]["change_summary"] if doc_rounds else ""
                )
            st.code(st.session_state[next_prompt_key], language="text")

            st.markdown("**다른 AI에게서 받은 답변을 아래에 붙여넣고 반영하세요**")
            external_input = st.text_area(
                "다른 AI로부터 받은 답변을 여기에 붙여넣으세요",
                height=200,
                key=f"{key_prefix}_external_ai_input_area_{current_round}",
            )
            if st.button(
                f"이 답변을 반영해서 {current_round + 1}차 확정 문서 만들기",
                key=f"{key_prefix}_apply_feedback_btn_{current_round}",
                type="primary",
            ):
                if not external_input.strip():
                    st.warning("붙여넣은 외부 AI 답변이 비어 있습니다.")
                else:
                    with st.spinner("외부 AI 의견을 반영해 문서를 다시 작성하는 중입니다... (웹검색 없이 빠르게 처리)"):
                        new_document, change_summary = engine.apply_external_feedback(
                            question=question,
                            current_document=current_document,
                            external_ai_input=external_input.strip(),
                        )

                    # 방금 보낸 질문과 그 답변을 직전 라운드 기록에 채워 넣음
                    doc_rounds[-1]["next_prompt"] = st.session_state.get(next_prompt_key, "")
                    doc_rounds[-1]["external_ai_input"] = external_input.strip()
                    doc_rounds[-1]["source"] = "manual"

                    next_round_no = current_round + 1
                    doc_rounds.append({
                        "round": next_round_no,
                        "document": new_document,
                        "change_summary": change_summary,
                        "verification_detail": "",
                        "next_prompt": "",
                        "external_ai_input": "",
                    })
                    st.session_state[rounds_key] = doc_rounds
                    st.session_state[current_doc_key] = new_document

                    # 다음 라운드용 질문을 곧바로 자동 생성
                    new_next_prompt = engine.build_next_round_prompt(question, new_document, change_summary)
                    st.session_state[next_prompt_key] = new_next_prompt
                    doc_rounds[-1]["next_prompt"] = new_next_prompt

                    # 안전장치 — 이번 라운드도 곧바로 백업
                    if engine.sheet_logger and engine.sheet_logger.enabled:
                        engine.sheet_logger.save_crosscheck_round(
                            session_id=crosscheck_session_id,
                            round_no=next_round_no,
                            question=question,
                            original_answer=answer,
                            verification_text=change_summary,
                            cross_prompt=new_next_prompt,
                            external_ai_input=external_input.strip(),
                            next_verification_text=new_document,
                        )
                    st.rerun()

        # ── 확정 (탭과 무관하게 항상 표시) ──────────────────────────
        st.divider()
        st.caption("위 과정(자동/수동 어느 쪽이든)은 원하는 만큼 반복할 수 있습니다. 충분히 다듬어졌다면 아래 버튼으로 확정하세요.")
        if st.button(
            f"이걸로 확정 (현재 {current_round}차 문서를 저장 대상으로 사용) →",
            key=f"{key_prefix}_proceed_btn",
            type="primary",
        ):
            # 이 시점부터는 기존 "검증대기" 탭 백업이 이어받으므로, 교차검증대기
            # 탭에 쌓아둔 라운드 기록은 더 이상 필요 없음. 정리함.
            if engine.sheet_logger and engine.sheet_logger.enabled:
                engine.sheet_logger.delete_crosscheck_session(crosscheck_session_id)
            st.session_state[proceed_key] = True
            st.rerun()
        return

    # ------------------------------------------------------------------
    # 확정 단계 — 이미 매 라운드 본문에 반영이 끝났으므로, 여기서는 추가
    # 자동 수정 없이 곧바로 파일선택/PIN 단계로 넘어감.
    # ------------------------------------------------------------------
    st.markdown("### ③ 저장할 내용 확인/수정")

    if st.button("↩ 라운드 진행 단계로 돌아가기 (추가로 교차검증하고 싶을 때)", key=f"{key_prefix}_back_to_verify_btn"):
        st.session_state[proceed_key] = False
        st.rerun()

    # 기본값을 마지막 라운드의 확정 문서로 설정. 매 라운드 이미 본문에 반영이
    # 끝났으므로, 별도의 "자동 수정" 호출이 더 필요하지 않음.
    if edited_key not in st.session_state:
        st.session_state[edited_key] = current_document

    st.caption("지금까지의 라운드에서 반영된 내용이 이미 아래에 담겨 있습니다. 한번 훑어보고 필요하면 추가로 고쳐주세요.")
    st.session_state[edited_key] = st.text_area(
        "지식베이스에 저장될 최종 내용",
        value=st.session_state[edited_key],
        height=500,
        key=f"{key_prefix}_edit_textarea",
    )

    # 파일 선택
    recommended_file = st.session_state.get(f"{key_prefix}_recommended_file", "")
    recommended_reason = st.session_state.get(f"{key_prefix}_recommended_reason", "")
    if recommended_reason:
        st.caption(f"추천 저장 파일: {recommended_file} — {recommended_reason}")
    default_idx = (
        engine.KNOWLEDGE_FILE_OPTIONS.index(recommended_file)
        if recommended_file in engine.KNOWLEDGE_FILE_OPTIONS
        else 0
    )
    target_file = st.selectbox(
        "④ 저장할 지식베이스 파일",
        options=engine.KNOWLEDGE_FILE_OPTIONS,
        index=default_idx,
        key=f"{key_prefix}_target_file_select",
    )

    # PIN 입력 후 최종 저장
    with st.form(f"{key_prefix}_final_confirm_form"):
        pin_input = st.text_input("⑤ 회계사 확정 PIN", type="password", key=f"{key_prefix}_final_pin_input")
        col_a, col_b = st.columns([1, 1])
        with col_a:
            submitted = st.form_submit_button("확정 저장 실행", type="primary", use_container_width=True)
        with col_b:
            cancelled = st.form_submit_button("취소", use_container_width=True)

        if submitted:
            if engine.verify_pin(pin_input):
                saved_path = engine.confirm_to_knowledge_base(
                    question=question,
                    confirmed_content=st.session_state[edited_key],
                    target_file=target_file,
                )

                # 종합문서 확정인 경우, 지식베이스에 새 항목을 추가하는 것과
                # 별개로 원본 '종합문서' 탭 행 자체도 확정된 최종본으로
                # 덮어씀 — 검증 전 원본과 확정본이 따로 남아 헷갈리지 않도록
                # 함(2026-07-09, 사용자 확인 후 결정된 방침).
                if source_summary_timestamp and engine.sheet_logger and engine.sheet_logger.enabled:
                    engine.sheet_logger.update_summary(
                        timestamp=source_summary_timestamp,
                        new_summary_text=st.session_state[edited_key],
                    )
                # 개별 질문/답변 확정인 경우: 내용은 그대로 두고 검색기록
                # 탭의 '확정여부'만 표시(종합문서와 달리 원본을 덮어쓰지 않음).
                if source_log_timestamp and engine.sheet_logger and engine.sheet_logger.enabled:
                    engine.sheet_logger.mark_log_confirmed(
                        timestamp=source_log_timestamp,
                        question=question,
                    )
                # 이번에 확정한 항목의 배지를 즉시 "✅ 확정됨"으로 바꾸기 위한
                # 세션 플래그. 화면에 종합문서를 표시하는 곳에서 참고함.
                st.session_state[f"{key_prefix}_kb_confirmed"] = True

                pending_ts = st.session_state.get(f"{key_prefix}_pending_ts")
                if pending_ts and engine.sheet_logger and engine.sheet_logger.enabled:
                    engine.sheet_logger.delete_pending_verification(pending_ts)
                st.session_state.pop(f"{key_prefix}_pending_ts", None)

                if engine.sheet_logger and engine.sheet_logger.enabled:
                    engine.sheet_logger.delete_crosscheck_session(crosscheck_session_id)

                st.session_state["kb_confirm_target"] = None
                st.session_state.pop(current_doc_key, None)
                st.session_state.pop(rounds_key, None)
                st.session_state.pop(next_prompt_key, None)
                st.session_state.pop(proceed_key, None)
                st.session_state.pop(edited_key, None)

                # 설계 의도 (복사 버튼 작업과 함께 수정): 기존에는 이 분기에
                # st.rerun()이 없어서, 위에서 kb_confirm_target 등 세션 상태는
                # 이미 정리됐는데도 화면(작업 공간, "새 주제 시작" 버튼 등)은
                # 이번 스크립트 실행분의 이전 상태 그대로 남아있어 "화면이
                # 멈춘 것처럼" 보였음(취소 버튼 쪽은 st.rerun()이 있어서
                # 정상적으로 화면이 정리됐던 것과 비대칭이었음). st.rerun()을
                # 걸면 이번에 만든 st.success 메시지가 그대로 사라지므로,
                # 저장 경로를 세션에 잠깐 담아뒀다가 리런 직후 화면 상단에서
                # 다시 보여주는 방식으로 처리함(아래 "저장 완료 배너" 참고).
                st.session_state["_kb_last_saved_path"] = saved_path
                st.rerun()
            else:
                st.error("PIN이 일치하지 않습니다.")
        if cancelled:
            if engine.sheet_logger and engine.sheet_logger.enabled:
                engine.sheet_logger.delete_crosscheck_session(crosscheck_session_id)
            st.session_state["kb_confirm_target"] = None
            st.session_state.pop(current_doc_key, None)
            st.session_state.pop(rounds_key, None)
            st.session_state.pop(next_prompt_key, None)
            st.session_state.pop(proceed_key, None)
            st.session_state.pop(edited_key, None)
            st.rerun()


@st.dialog("검색 기록 상세보기", width="large")
def show_log_dialog():
    """
    사이드바 '검색 기록'에서 항목을 클릭하면 뜨는 모달 팝업.
    전체 질문/답변을 보여주고, 그 자리에서 바로 파일 저장 및 지식베이스 확정도 가능.
    (현재 대화 화면의 개별 답변 카드에 있는 기능과 동일한 동작을 과거 기록에 대해서도 제공)
    """
    row = st.session_state.get("_dialog_log_row")
    if not row:
        st.write("표시할 기록이 없습니다.")
        return

    question = row.get("질문", "")
    answer = row.get("전체답변", row.get("답변요약", "(내용 없음)"))

    st.caption(f"{row.get('일시', '')}  ·  {row.get('사용자구분', '')}")
    st.markdown(f"**질문**")
    st.write(question)
    dialog_key_base = f"dialog_{row.get('일시', '')}".replace(" ", "_").replace(":", "")
    if row.get("확정여부") == "확정됨" or st.session_state.get(f"{dialog_key_base}_kb_confirmed"):
        st.success("✅ 지식베이스에 확정된 질문/답변입니다.")
    st.divider()
    render_copyable_text(answer, key=f"{dialog_key_base}_answer")
    st.divider()

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col1:
        save_show_key = f"{dialog_key_base}_save_show"
        if st.button("이 답변 저장", key=f"{dialog_key_base}_save"):
            st.session_state[save_show_key] = True
        if st.session_state.get(save_show_key):
            offer_downloads(
                title="세무질의 자문 기록",
                question=question,
                answer=answer,
                filename_hint=f"자문_{row.get('일시', '').replace('-', '').replace(':', '').replace(' ', '_')}",
                key_prefix=dialog_key_base,
            )

    with col2:
        if is_admin:
            render_confirm_to_kb_button(
                question=question, answer=answer, key_prefix=dialog_key_base,
                dialog_row_key="_dialog_log_row",
                source_log_timestamp=row.get("일시", ""),
                already_confirmed=(
                    row.get("확정여부") == "확정됨"
                    or bool(st.session_state.get(f"{dialog_key_base}_kb_confirmed"))
                ),
            )

    with col3:
        # 검색 기록 삭제 — 되돌릴 수 없는 작업이므로 지식베이스 확정과 동일하게
        # 회계사 확정 PIN을 입력해야만 실행되도록 보호함.
        delete_confirm_key = f"{dialog_key_base}_show_delete_confirm"
        if st.button("이 기록 삭제", key=f"{dialog_key_base}_delete_btn"):
            st.session_state[delete_confirm_key] = True

        if st.session_state.get(delete_confirm_key):
            if not engine.has_pin_set():
                st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
            else:
                with st.form(f"{dialog_key_base}_delete_form"):
                    st.caption("삭제하면 구글 시트에서 이 행이 완전히 사라지며, 되돌릴 수 없습니다.")
                    delete_pin_input = st.text_input(
                        "회계사 확정 PIN", type="password", key=f"{dialog_key_base}_delete_pin_input"
                    )
                    if st.form_submit_button("삭제 실행"):
                        if engine.verify_pin(delete_pin_input):
                            if not (engine.sheet_logger and engine.sheet_logger.enabled):
                                st.error("구글 시트 로깅이 비활성 상태라 삭제할 수 없습니다.")
                            else:
                                deleted = engine.sheet_logger.delete_log(
                                    timestamp=row.get("일시", ""), question=question
                                )
                                if deleted:
                                    st.success("기록이 삭제되었습니다.")
                                    st.session_state[delete_confirm_key] = False
                                    # 사이드바에 캐시된 목록도 갱신해야 화면에서 사라짐
                                    st.session_state["_loaded_logs"] = [
                                        r for r in st.session_state.get("_loaded_logs", [])
                                        if not (r.get("일시") == row.get("일시") and r.get("질문") == question)
                                    ]
                                    st.rerun()
                                else:
                                    st.error("삭제에 실패했습니다. 해당 행을 찾지 못했거나 시트 접근에 문제가 있습니다.")
                        else:
                            st.error("PIN이 일치하지 않습니다.")


@st.dialog("종합 문서 기록 상세보기", width="large")
def show_summary_log_dialog():
    """
    사이드바 '종합 문서 기록'에서 항목을 클릭하면 뜨는 모달 팝업.
    종합 문서 전체 내용을 보여주고, 다시 파일로 저장할 수 있음.

    설계 변경 (종합문서 지식베이스 확정 기능 추가와 함께):
    - 예전에는 "종합 문서는 여러 질문을 재구성한 2차 가공물이므로 지식베이스
      확정 대상에서는 제외"하는 방침이었으나, 이후 오히려 "꼬리질문으로 정정된
      최종 결론이 반영된 종합 문서가 개별 답변보다 지식베이스 확정에 더
      적합하다"는 방향으로 바뀌었음. 현재 대화 화면의 종합 문서에는 이미 확정
      버튼이 붙어 있는데, 과거 검색 기록(구글시트 '종합문서' 탭)에서 불러온
      종합 문서에는 빠져 있어 비대칭이었던 것을 여기서 맞춤.
    - show_log_dialog(검색 기록 상세보기)와 동일한 패턴으로
      render_confirm_to_kb_button을 추가하고, dialog_row_key를 넘겨 다이얼로그가
      안전하게 닫히고 메인 화면 작업 공간으로 넘어가도록 함.
    """
    row = st.session_state.get("_dialog_summary_row")
    if not row:
        st.write("표시할 기록이 없습니다.")
        return

    summary_text = row.get("종합문서전체", row.get("종합문서요약", "(내용 없음)"))

    st.caption(
        f"{row.get('일시', '')}  ·  포함된 질의 {row.get('포함된질의건수', '?')}건  ·  "
        f"{row.get('사용자구분', '')}"
    )
    dialog_key_base = f"sdialog_{row.get('일시', '')}".replace(" ", "_").replace(":", "")
    # 구글시트 '확정여부' 컬럼을 그대로 신뢰함 — 세션 상태가 아니라 시트 값을
    # 기준으로 판단하므로, 다른 브라우저/다른 사람이 확정한 경우에도 정확함.
    if row.get("확정여부") == "확정됨" or st.session_state.get(f"{dialog_key_base}_kb_confirmed"):
        st.success("✅ 지식베이스에 확정된 문서입니다.")
    else:
        st.warning("⚠️ 아직 지식베이스에 확정되지 않은 AI 생성 문서입니다. 검증 절차를 거치지 않아 오류가 있을 수 있습니다.")
    render_copyable_text(summary_text, key=f"{dialog_key_base}_summary")
    st.divider()

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col1:
        save_show_key = f"{dialog_key_base}_save_show"
        if st.button("이 종합 문서 다시 저장", key=f"{dialog_key_base}_save"):
            st.session_state[save_show_key] = True
        if st.session_state.get(save_show_key):
            offer_downloads(
                title="세무질의 종합 자문 문서",
                question="",
                answer=summary_text,
                filename_hint=f"종합자문_{row.get('일시', '').replace('-', '').replace(':', '').replace(' ', '_')}",
                key_prefix=dialog_key_base,
            )

    with col2:
        if is_admin:
            render_confirm_to_kb_button(
                question=(
                    f"[종합문서] {row.get('일시', '')} 종합 "
                    f"(포함된 질의 {row.get('포함된질의건수', '?')}건)"
                ),
                answer=summary_text,
                key_prefix=dialog_key_base,
                dialog_row_key="_dialog_summary_row",
                source_summary_timestamp=row.get("일시", ""),
                already_confirmed=(
                    row.get("확정여부") == "확정됨"
                    or bool(st.session_state.get(f"{dialog_key_base}_kb_confirmed"))
                ),
            )

    with col3:
        # 종합 문서 삭제 — 검색 기록 삭제와 동일하게 회계사 확정 PIN으로 보호함.
        delete_confirm_key = f"{dialog_key_base}_show_delete_confirm"
        if st.button("이 종합 문서 삭제", key=f"{dialog_key_base}_delete_btn"):
            st.session_state[delete_confirm_key] = True

        if st.session_state.get(delete_confirm_key):
            if not engine.has_pin_set():
                st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
            else:
                with st.form(f"{dialog_key_base}_delete_form"):
                    st.caption("삭제하면 구글 시트에서 이 행이 완전히 사라지며, 되돌릴 수 없습니다.")
                    delete_pin_input = st.text_input(
                        "회계사 확정 PIN", type="password", key=f"{dialog_key_base}_delete_pin_input"
                    )
                    if st.form_submit_button("삭제 실행"):
                        if engine.verify_pin(delete_pin_input):
                            if not (engine.sheet_logger and engine.sheet_logger.enabled):
                                st.error("구글 시트 로깅이 비활성 상태라 삭제할 수 없습니다.")
                            else:
                                deleted = engine.sheet_logger.delete_summary(timestamp=row.get("일시", ""))
                                if deleted:
                                    st.success("종합 문서 기록이 삭제되었습니다.")
                                    st.session_state[delete_confirm_key] = False
                                    st.session_state["_loaded_summaries"] = [
                                        r for r in st.session_state.get("_loaded_summaries", [])
                                        if r.get("일시") != row.get("일시")
                                    ]
                                    st.rerun()
                                else:
                                    st.error("삭제에 실패했습니다. 해당 행을 찾지 못했거나 시트 접근에 문제가 있습니다.")
                        else:
                            st.error("PIN이 일치하지 않습니다.")


# ----------------------------------------------------------------------
# 사이드바: 상태 및 옵션
# ----------------------------------------------------------------------
with st.sidebar:
    st.subheader("시스템 상태")
    st.write(f"지식 베이스 경로: `{engine.knowledge_dir}`")
    st.write(f"법제처 연동: {'사용 중' if engine.law_client else '미사용 (.env에 LAW_API_OC 없음)'}")
    nts_enabled = os.getenv("ENABLE_NTS_SEARCH", "").strip().lower() in ("true", "1", "yes")
    st.write(f"국세청 검색 보강: {'사용 중' if nts_enabled else '미사용'}")

    st.divider()
    st.subheader("로그인 정보")
    st.write(f"현재 역할: **{'관리자 (회계사)' if is_admin else '직원'}**")
    if st.button("로그아웃", key="logout_btn"):
        st.session_state.app_role = None
        st.rerun()

    if is_admin:
        st.divider()
        st.subheader("확정 PIN 관리")
        st.caption("AI 답변을 지식베이스에 확정 저장할 때 필요한 PIN입니다. 회계사만 알아야 합니다.")

        if not engine.has_pin_set():
            st.info("아직 PIN이 설정되지 않았습니다.")
            with st.form("set_pin_form", clear_on_submit=True):
                new_pin = st.text_input("새 PIN 설정 (4자 이상)", type="password")
                new_pin_confirm = st.text_input("PIN 확인", type="password")
                if st.form_submit_button("PIN 설정"):
                    if len(new_pin.strip()) < 4:
                        st.error("PIN은 4자 이상으로 설정해주세요.")
                    elif new_pin != new_pin_confirm:
                        st.error("입력한 PIN이 서로 다릅니다.")
                    else:
                        engine.set_pin(new_pin)
                        st.success("PIN이 설정되었습니다.")
                        st.rerun()
        else:
            with st.expander("PIN 변경", expanded=False):
                with st.form("change_pin_form", clear_on_submit=True):
                    current_pin = st.text_input("현재 PIN", type="password")
                    new_pin = st.text_input("새 PIN (4자 이상)", type="password")
                    new_pin_confirm = st.text_input("새 PIN 확인", type="password")
                    if st.form_submit_button("PIN 변경"):
                        if not engine.verify_pin(current_pin):
                            st.error("현재 PIN이 일치하지 않습니다.")
                        elif len(new_pin.strip()) < 4:
                            st.error("새 PIN은 4자 이상으로 설정해주세요.")
                        elif new_pin != new_pin_confirm:
                            st.error("입력한 새 PIN이 서로 다릅니다.")
                        else:
                            engine.set_pin(new_pin)
                            st.success("PIN이 변경되었습니다.")
                            st.rerun()

        # ------------------------------------------------------------------
        # 로그인 비밀번호 관리 (2026-06-27 추가)
        # ------------------------------------------------------------------
        # 관리자/직원 로그인 비밀번호를 .env/Secrets가 아니라 여기서 직접
        # 변경할 수 있게 함(PIN 변경과 동일한 패턴 — 현재 비밀번호 확인 후 변경).
        # 구글시트 연동이 비활성 상태면 이 기능 자체가 의미 없으므로(환경변수
        # 방식으로 폴백 중이라는 뜻) 숨김.
        if engine.sheet_logger and engine.sheet_logger.enabled:
            with st.expander("로그인 비밀번호 관리", expanded=False):
                st.caption("관리자/직원이 로그인할 때 쓰는 비밀번호입니다. 변경 시 즉시 적용됩니다.")
                pw_role_label = st.radio(
                    "변경할 대상", options=["관리자(회계사)", "직원"], horizontal=True, key="pw_change_role"
                )
                pw_role = "admin" if pw_role_label == "관리자(회계사)" else "staff"
                with st.form("change_login_pw_form", clear_on_submit=True):
                    current_admin_pw = st.text_input(
                        "현재 관리자(회계사) 로그인 비밀번호",
                        type="password",
                        help="지식베이스 확정용 PIN이 아니라, 로그인 화면에서 입력하는 "
                        "관리자 비밀번호입니다. 실수로 잘못 바꾸는 것을 막기 위한 본인 확인입니다.",
                    )
                    new_login_pw = st.text_input("새 비밀번호 (4자 이상)", type="password")
                    new_login_pw_confirm = st.text_input("새 비밀번호 확인", type="password")
                    if st.form_submit_button("비밀번호 변경"):
                        # 비상 복구 비밀번호로 로그인한 경우, 원래 비밀번호를
                        # 몰라서 여기 온 것이므로 본인확인 단계에서도 비상
                        # 복구 비밀번호를 그대로 인정함(그래야 실제로 복구가
                        # 끝까지 됨 — 안 그러면 "잊어버린 비밀번호를 다시
                        # 입력하라"는 막다른 길이 됨).
                        current_ok = (
                            engine.sheet_logger.verify_account_password("admin", current_admin_pw)
                            or current_admin_pw.strip() == _get_master_reset_password()
                        )
                        if not current_ok:
                            st.error("현재 관리자(회계사) 비밀번호가 일치하지 않습니다.")
                        elif len(new_login_pw.strip()) < 4:
                            st.error("새 비밀번호는 4자 이상으로 설정해주세요.")
                        elif new_login_pw != new_login_pw_confirm:
                            st.error("입력한 새 비밀번호가 서로 다릅니다.")
                        else:
                            engine.sheet_logger.set_account_password(pw_role, new_login_pw)
                            st.session_state["_logged_in_via_master_reset"] = False
                            st.success(f"{pw_role_label} 비밀번호가 변경되었습니다.")

        if st.button(
            "지식 베이스 새로고침",
            help="로컬 _knowledge 폴더 파일을 수정했거나, 다른 곳에서 확정 저장한 "
            "구글시트 지식베이스 내용을 즉시 반영하고 싶을 때 누르세요.",
        ):
            engine.load_knowledge_base(force_reload=True)
            st.success("지식 베이스를 다시 불러왔습니다.")

        # ------------------------------------------------------------------
        # 검증대기 불러오기 (2026-06-27 추가)
        # ------------------------------------------------------------------
        # 설계 의도: 화면이 튕겨서(WebSocket 연결 끊김 등) 검증까지 끝낸 작업을
        # 잃어버렸을 때, 웹검색을 처음부터 다시 돌리지 않고 여기서 즉시 복원해
        # PIN 입력 단계로 바로 이어갈 수 있게 함. PIN 관리·확정 저장과 같은
        # 관리자 전용 흐름의 일부이므로 is_admin으로 같이 감쌈.
        if engine.sheet_logger and engine.sheet_logger.enabled:
            pending_list = engine.sheet_logger.list_pending_verifications()
            if pending_list:
                st.divider()
                st.subheader("검증대기 불러오기")
                st.caption(
                    f"화면이 튕겨도 잃지 않도록 자동 백업된 검증 결과입니다 "
                    f"({len(pending_list)}건). 불러오면 웹검색을 다시 돌리지 않고 "
                    f"PIN 입력 단계로 바로 이동합니다."
                )
                for p_row in pending_list:
                    p_ts = p_row.get("일시", "")
                    p_question = p_row.get("질문", "")
                    short_q = p_question[:30] + ("..." if len(p_question) > 30 else "")
                    p_key_base = f"pending_{p_ts}".replace(" ", "_").replace(":", "").replace("-", "")

                    col_load, col_del = st.columns([5, 1])
                    with col_load:
                        if st.button(f"📋 {p_ts} — {short_q}", key=f"load_{p_key_base}", use_container_width=True):
                            restore_key_prefix = p_key_base
                            st.session_state["kb_confirm_target"] = {
                                "question": p_question,
                                "answer": p_row.get("원본답변", ""),
                                "key_prefix": restore_key_prefix,
                            }
                            # v1.7 — 검증대기 복원은 이미 검증+자동수정이 끝난 결과를
                            # 그대로 가져오는 것이므로, "1라운드" 기록으로 채워두고
                            # 곧바로 ③ 결과 확인 단계로 건너뛰게 함.
                            restored_doc = p_row.get("수정된내용", "")
                            st.session_state[f"{restore_key_prefix}_current_doc"] = restored_doc
                            st.session_state[f"{restore_key_prefix}_doc_rounds"] = [{
                                "round": 1,
                                "document": restored_doc,
                                "change_summary": p_row.get("수정요약", ""),
                                "verification_detail": "(검증대기에서 복원됨 — 원래 검증 상세 내용은 표시되지 않습니다.)",
                                "next_prompt": "",
                                "external_ai_input": "",
                            }]
                            st.session_state[f"{restore_key_prefix}_recommended_file"] = p_row.get("추천파일", "")
                            st.session_state[f"{restore_key_prefix}_recommended_reason"] = p_row.get("추천이유", "")
                            st.session_state[f"{restore_key_prefix}_edited_content"] = restored_doc
                            st.session_state[f"{restore_key_prefix}_pending_ts"] = p_ts
                            st.session_state[f"{restore_key_prefix}_proceed_to_confirm"] = True
                            st.info("화면 맨 아래 '지식베이스 확정 저장 작업 공간'으로 이동해 진행해주세요.")
                            st.rerun()

                    # 삭제 — 오래돼서 더 진행할 필요 없는 항목을 구글시트에
                    # 직접 들어가지 않고도 여기서 바로 정리할 수 있게 함
                    # (2026-07-02 추가). 실수로 한 번에 지워지지 않도록 클릭
                    # 시 확인 버튼을 한 번 더 거치는 2단계 확인 방식 사용.
                    del_confirm_key = f"del_confirm_{p_key_base}"
                    with col_del:
                        if st.button("🗑", key=f"del_{p_key_base}", help="이 검증대기 항목 삭제"):
                            st.session_state[del_confirm_key] = True
                    if st.session_state.get(del_confirm_key):
                        st.caption(f"'{short_q}' 항목을 삭제할까요? 되돌릴 수 없습니다.")
                        c_yes, c_no = st.columns(2)
                        with c_yes:
                            if st.button("삭제 확정", key=f"del_yes_{p_key_base}", use_container_width=True):
                                engine.sheet_logger.delete_pending_verification(p_ts)
                                st.session_state[del_confirm_key] = False
                                st.rerun()
                        with c_no:
                            if st.button("취소", key=f"del_no_{p_key_base}", use_container_width=True):
                                st.session_state[del_confirm_key] = False
                                st.rerun()

        # ------------------------------------------------------------------
        # 교차검증대기 불러오기 (2026-06-28 추가)
        # ------------------------------------------------------------------
        # 설계 의도: 위의 "검증대기"는 ①②③ 자동검증+수정이 끝난 "확정 직전"
        # 단계만 백업함. 이쪽은 그보다 앞선 "②번 교차검증 스레드"(1차 검증 →
        # 다른 AI에게 질문 → 답변 받아 재검증 → 2차, 3차...) 도중에 화면이
        # 튕긴 경우를 위한 것임. 회계사가 다른 AI 사이트로 탭을 옮겨 한참
        # 머무는 구간이라 WebSocket이 가장 자주 끊기는 지점인데, 여기서
        # 끊기면 진행하던 모든 라운드가 사라지던 문제를 막기 위함.
        if engine.sheet_logger and engine.sheet_logger.enabled:
            crosscheck_sessions = engine.sheet_logger.list_crosscheck_sessions()
            if crosscheck_sessions:
                st.divider()
                st.subheader("교차검증 진행 중 항목 불러오기")
                st.caption(
                    f"다른 AI에게 교차검증을 요청하는 도중 화면이 튕겨도 잃지 않도록 "
                    f"자동 백업된 진행 기록입니다 ({len(crosscheck_sessions)}건). 불러오면 "
                    f"지금까지 진행했던 라운드 그대로 교차검증을 이어갈 수 있습니다."
                )
                for cc_session in crosscheck_sessions:
                    cc_sid = cc_session["session_id"]
                    cc_q = cc_session["question"]
                    cc_short_q = cc_q[:30] + ("..." if len(cc_q) > 30 else "")
                    cc_label = f"🔄 {cc_session['updated_at']} — {cc_short_q} ({cc_session['latest_round']}차까지 진행)"
                    cc_key_base = f"crosscheck_{cc_sid}".replace("-", "")

                    col_load, col_del = st.columns([5, 1])
                    with col_load:
                        if st.button(cc_label, key=f"load_{cc_key_base}", use_container_width=True):
                            cc_rounds = engine.sheet_logger.get_crosscheck_rounds(cc_sid)
                            if not cc_rounds:
                                st.warning("이 항목의 라운드 기록을 찾지 못했습니다.")
                            else:
                                restore_key_prefix = cc_key_base
                                original_answer = cc_rounds[0]["original_answer"]
                                last_round = cc_rounds[-1]

                                st.session_state["kb_confirm_target"] = {
                                    "question": cc_q,
                                    "answer": original_answer,
                                    "key_prefix": restore_key_prefix,
                                    # 기존 세션ID를 그대로 재사용함 — 새로 만들면
                                    # 이어서 백업되는 라운드가 다른 세션으로 갈라져
                                    # 기존 구글시트 기록과 끊어짐.
                                    "crosscheck_session_id": cc_sid,
                                }
                                # v1.7 — 저장 시점에 verification_text 컬럼에는
                                # change_summary가, next_verification_text 컬럼에는
                                # 그 라운드의 확정 문서 전체가 들어가 있음(엔진의
                                # save_crosscheck_round 호출부 참고). 그대로 매핑해
                                # doc_rounds 구조로 복원함.
                                st.session_state[f"{restore_key_prefix}_doc_rounds"] = [
                                    {
                                        "round": r["round"],
                                        "document": r.get("next_verification_text", ""),
                                        "change_summary": r.get("verification_text", ""),
                                        "verification_detail": "",
                                        "next_prompt": r.get("cross_prompt", ""),
                                        "external_ai_input": r.get("external_ai_input", ""),
                                    }
                                    for r in cc_rounds
                                ]
                                # 마지막 라운드의 확정 문서를 "지금 진행 중인" 문서로 이어받음.
                                resumed_document = (
                                    last_round.get("next_verification_text") or original_answer
                                )
                                st.session_state[f"{restore_key_prefix}_current_doc"] = resumed_document
                                st.session_state[f"{restore_key_prefix}_next_prompt"] = engine.build_next_round_prompt(
                                    cc_q, resumed_document, last_round.get("verification_text", "")
                                )
                                st.info("화면 맨 아래 '지식베이스 확정 저장 작업 공간'으로 이동해 이어서 진행해주세요.")
                                st.rerun()

                    # 삭제 — 검증대기와 동일한 2단계 확인 방식 (2026-07-02 추가)
                    cc_del_confirm_key = f"del_confirm_{cc_key_base}"
                    with col_del:
                        if st.button("🗑", key=f"del_{cc_key_base}", help="이 교차검증 진행 항목 삭제"):
                            st.session_state[cc_del_confirm_key] = True
                    if st.session_state.get(cc_del_confirm_key):
                        st.caption(f"'{cc_short_q}' 항목을 삭제할까요? 진행 중이던 라운드 기록이 모두 사라지며 되돌릴 수 없습니다.")
                        c_yes, c_no = st.columns(2)
                        with c_yes:
                            if st.button("삭제 확정", key=f"del_yes_{cc_key_base}", use_container_width=True):
                                engine.sheet_logger.delete_crosscheck_session(cc_sid)
                                st.session_state[cc_del_confirm_key] = False
                                st.rerun()
                        with c_no:
                            if st.button("취소", key=f"del_no_{cc_key_base}", use_container_width=True):
                                st.session_state[cc_del_confirm_key] = False
                                st.rerun()

    st.divider()
    st.subheader("검색 기록")
    st.caption("개별 질문 하나하나에 대한 답변 기록입니다.")
    if engine.sheet_logger and engine.sheet_logger.enabled:
        st.success("구글 시트 로깅 사용 중")
        n_logs = st.slider("불러올 최근 기록 수", min_value=10, max_value=200, value=30, step=10)
        search_term = st.text_input("질문 내용 검색 (비워두면 전체 표시)", key="log_search_term")

        if st.button("기록 불러오기", key="load_logs_btn"):
            with st.spinner("구글 시트에서 기록을 불러오는 중..."):
                recent = engine.sheet_logger.get_recent_logs(n_logs)
            st.session_state["_loaded_logs"] = recent

        loaded = st.session_state.get("_loaded_logs", [])
        if loaded:
            filtered = loaded
            if search_term.strip():
                filtered = [
                    row for row in loaded
                    if search_term.strip() in row.get("질문", "")
                ]
            st.caption(f"{len(filtered)}건 표시 (전체 불러온 기록 {len(loaded)}건 중)")

            for idx, row in enumerate(filtered):
                badge = "✅ " if row.get("확정여부") == "확정됨" else ""
                label = f"{badge}{row.get('일시', '')} | {row.get('사용자구분', '')} | {row.get('질문', '')[:30]}"
                if st.button(label, key=f"log_open_{idx}", use_container_width=True):
                    st.session_state["_dialog_log_row"] = row
                    show_log_dialog()
        else:
            st.write("위 '기록 불러오기' 버튼을 눌러 구글 시트에서 기록을 가져오세요.")
    else:
        # enabled=False인 두 가지 경우를 구분해서 보여줌:
        # (1) 아예 설정을 안 한 경우 (error_message가 비어있음) → 안내문만 표시
        # (2) 설정을 시도했는데 실패한 경우 (error_message가 있음) → 실제 실패 원인을 그대로 보여줌
        #     (이전에는 두 경우 모두 같은 안내문으로 뭉뚱그려져서, 실제로 무엇이 잘못됐는지
        #     화면만 보고는 알 수 없었음. 진단을 쉽게 하기 위해 분리함.)
        error_msg = engine.sheet_logger.error_message if engine.sheet_logger else ""
        if error_msg:
            st.error(f"구글 시트 로깅 설정을 시도했지만 실패했습니다.\n\n원인: {error_msg}")
        else:
            st.info(
                "구글 시트 로깅 미사용.\n\n"
                "- 로컬 환경: `.env`에 GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_PATH를 설정하세요.\n"
                "- 웹 배포 환경: Streamlit Secrets에 GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_JSON을 설정하세요.\n\n"
                "설정하면 모든 질의응답이 구글 스프레드시트에 자동 기록됩니다."
            )

    st.divider()
    st.subheader("종합 문서 기록")
    st.caption(
        "여러 질문을 묶어 재구성한 종합 문서 기록입니다. "
        "'종합 문서 생성' 시점에 자동으로 여기에 남습니다 "
        "(개별 질문 기록과는 별도 탭에 저장되어 섞이지 않습니다). "
        "✅ 확정됨 표시가 없으면 검증 절차를 거치지 않은 AI 생성 문서이니 참고하세요."
    )
    if engine.sheet_logger and engine.sheet_logger.enabled:
        n_summaries = st.slider(
            "불러올 최근 종합 문서 수", min_value=5, max_value=100, value=20, step=5, key="n_summaries_slider"
        )

        if st.button("종합 문서 기록 불러오기", key="load_summaries_btn"):
            with st.spinner("구글 시트에서 종합 문서 기록을 불러오는 중..."):
                recent_summaries = engine.sheet_logger.get_recent_summaries(n_summaries)
            st.session_state["_loaded_summaries"] = recent_summaries

        loaded_summaries = st.session_state.get("_loaded_summaries", [])
        if loaded_summaries:
            st.caption(f"{len(loaded_summaries)}건 표시")
            for idx, row in enumerate(loaded_summaries):
                badge = "✅ " if row.get("확정여부") == "확정됨" else "⚠️ "
                label = (
                    f"{badge}{row.get('일시', '')} | 질의 {row.get('포함된질의건수', '?')}건 | "
                    f"{row.get('종합문서요약', '')[:30]}"
                )
                if st.button(label, key=f"summary_open_{idx}", use_container_width=True):
                    st.session_state["_dialog_summary_row"] = row
                    show_summary_log_dialog()
        else:
            st.write("위 '종합 문서 기록 불러오기' 버튼을 눌러 구글 시트에서 가져오세요.")
    else:
        st.info("구글 시트 로깅 미사용 상태라 종합 문서 기록도 사용할 수 없습니다.")

    st.divider()
    with st.expander("진단 정보 (문제 발생 시 확인용)", expanded=False):
        # 구글 시트 관련 환경변수 진단
        # 주의: 실제 키 값은 절대 화면에 표시하지 않음. 존재 여부와 길이만 표시해서
        # "os.getenv()가 값을 제대로 읽고 있는지"만 확인할 수 있게 함.
        _gsheet_id = os.getenv("GOOGLE_SHEET_ID", "")
        _gcred_json = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
        _gcred_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "")
        _err = engine.sheet_logger.error_message if engine.sheet_logger else "(sheet_logger 없음)"
        if _err == "__NOT_CONFIGURED__":
            _err = "(설정 자체가 없음 — 정상. 환경변수가 비어있어서 시도조차 안 함)"
        elif _err == "":
            _err = "(빈 문자열 — try 블록 진입 후 self.enabled=True 직전에 멈췄을 가능성)"
        _tab_errors = getattr(engine.sheet_logger, "tab_errors", {}) if engine.sheet_logger else {}
        st.code(
            f"[구글 시트 관련]\n"
            f"GOOGLE_SHEET_ID 읽힘: {'예 (' + str(len(_gsheet_id)) + '자)' if _gsheet_id else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_JSON 읽힘: {'예 (' + str(len(_gcred_json)) + '자)' if _gcred_json else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_PATH 읽힘: {'예 (' + str(len(_gcred_path)) + '자)' if _gcred_path else '아니오 (빈 값)'}\n"
            f"sheet_logger.enabled: {engine.sheet_logger.enabled if engine.sheet_logger else '(sheet_logger 없음)'}\n"
            f"sheet_logger.error_message: {_err}\n"
            f"tab_errors (탭별 개별 오류, 있으면 그 탭 기능만 비활성): "
            f"{_tab_errors if _tab_errors else '(없음 — 모든 탭 정상)'}\n"
            f"\n[법제처 관련]\n"
            f"ENABLE_NTS_SEARCH 원본값: {os.getenv('ENABLE_NTS_SEARCH', '(없음)')!r}\n"
            f"LAW_API_OC 설정 여부: {'설정됨' if os.getenv('LAW_API_OC', '').strip() else '미설정'}\n"
            f"law_client 객체: {engine.law_client}\n"
            f"law_client 모듈 파일 경로: {getattr(__import__('law_api'), '__file__', '확인불가')}",
            language="text",
        )
        if st.button("국세청 검색 직접 테스트"):
            test_q = "음식점업을 영위하는 사업자가 면세 재화를 매입한 경우 매입세액공제는?"
            st.write(f"테스트 질문: {test_q}")

            # 1단계: 키워드 추출 결과를 그대로 노출 (repr로 공백/특수문자까지 확인 가능)
            extracted = engine._extract_search_keywords(test_q, max_keywords=3)
            st.write("추출된 키워드 (raw repr):")
            st.code(repr(extracted), language="python")

            # 2단계: 키워드 각각으로 직접 법제처 검색 호출 (engine을 거치지 않고 law_client 직접 사용)
            if extracted:
                import urllib.parse as _urlparse
                for kw in extracted:
                    try:
                        # 원시 URL과 원시 응답까지 직접 확인 (어디서 0건이 되는지 추적)
                        raw_url = "https://www.law.go.kr/DRF/lawSearch.do?" + _urlparse.urlencode({
                            "OC": engine.law_client.oc_key,
                            "target": "ntsCgmExpc",
                            "type": "XML",
                            "query": kw,
                            "display": 3,
                            "search": 2,
                        })
                        st.code(raw_url, language="text")
                        raw_response = engine.law_client._fetch(raw_url)
                        st.text_area(f"'{kw}' 원시 응답 (앞 1000자)", raw_response[:1000], height=150, key=f"raw_{kw}")

                        kw_result = engine.law_client.search_nts_interpretations(kw, display=3)
                        st.write(f"  - 키워드 {repr(kw)} → {len(kw_result)}건")
                        if kw_result:
                            st.code(str(kw_result[0]), language="python")
                    except Exception as e:
                        st.error(f"키워드 {repr(kw)} → 예외 발생: {type(e).__name__}: {e}")
            else:
                st.warning("키워드 추출 결과가 비어 있습니다 (Gemini 호출 자체가 실패했을 가능성)")

            # 3단계: 전체 파이프라인 결과
            test_result = engine.search_nts_interpretations(test_q)
            st.write("전체 파이프라인 결과 길이:", len(test_result))
            if test_result:
                st.text(test_result[:500])

    st.divider()
    st.subheader("저장 옵션")
    st.caption(
        "결과는 파일로 만들어 다운로드 버튼으로 받습니다 "
        "(서버에 저장되지 않으며, 받은 파일은 사용자 컴퓨터에만 남습니다)."
    )
    # 설계 의도 (2026-07-02 버그 수정): 기존에는 이 값들을 지역 변수로만
    # 두고 offer_downloads()가 그 이름을 그대로 가져다 썼음. 사이드바 코드가
    # 이 지점(스크립트의 뒷부분)까지 먼저 실행돼야만 값이 "존재"하는데,
    # st.dialog로 띄우는 팝업(검색 기록 상세보기 등)은 호출되는 순간
    # 그 아래 스크립트 실행을 멈춰버리는 특성이 있어서, 사이드바 위쪽의
    # "검색 기록" 항목을 눌러 다이얼로그가 뜨면 이 아래 코드가 아직 실행 전
    # 상태 → 다이얼로그 안에서 저장을 누르면 NameError가 났음. key=로
    # session_state에 연결해두면 실행 순서와 무관하게 항상 값을 읽을 수 있음.
    # 버그 수정 (2026-07-02): 파일 상단에서 이미 st.session_state.setdefault로
    # 이 키들의 초기값을 넣어뒀는데, 위젯 생성 시 default=/value=까지 같이
    # 주면 Streamlit이 "Session State로 값이 설정된 위젯에 default까지 줬다"는
    # 경고(policy 위반)를 띄움. 동작에는 지장 없지만 로그가 지저분해지므로,
    # 초기값 지정은 session_state.setdefault 쪽에만 맡기고 위젯에서는 뺌.
    save_formats = st.multiselect(
        "다운로드 형식 (여러 개 선택 가능)",
        options=["md", "docx", "pdf"],
        help="pdf는 서버에 한글 폰트가 없으면 생성되지 않을 수 있습니다. "
        "이 경우 md 또는 docx로 받아주세요.",
        key="save_formats",
    )
    use_custom_filename = st.checkbox("파일명 직접 지정", key="use_custom_filename")
    custom_filename = ""
    if use_custom_filename:
        custom_filename = st.text_input("파일명 (확장자 제외)", key="custom_filename")

    st.divider()
    st.subheader("전체 초기화")
    if st.button("모든 기록 삭제 (백로그 포함)"):
        st.session_state.current_thread = []
        st.session_state.backlog = []
        st.session_state.summary_doc = None
        st.rerun()


# ----------------------------------------------------------------------
# 저장 헬퍼
# ----------------------------------------------------------------------
def build_combined_text(turns):
    blocks = []
    for i, qa in enumerate(turns, start=1):
        blocks.append(f"[질의 {i}] {qa['question']}\n[회신 {i}]\n{qa['answer']}")
    return "\n\n---\n\n".join(blocks)


def run_summary(turns):
    combined_text = build_combined_text(turns)
    summary_prompt = f"""다음은 한 세션 동안 기장 직원이 순서대로 질문하고 AI가 답변한 여러 건의 세무 질의응답입니다.
이 내용을 종합하여, 회계사가 검토하기 편하도록 하나의 완성된 세무 자문 문서로 재구성하세요.

[작성 규칙]
1. 중복되거나 서로 연관된 질문/답변은 통합하여 정리하세요.
2. 전체 내용을 아우르는 "종합 요약"을 맨 앞에 추가하세요.
3. 이후 각 질의별 결론을 순서대로 정리하세요.
4. 추측하지 말고, 각 답변에 있던 내용만 사용하여 재구성하세요.
5. 이모지는 사용하지 마세요.
6. 마크다운 형식으로 작성하세요.

[질의응답 원문]
{combined_text}
"""
    response = engine.client.models.generate_content(
        model=engine.model_name,
        contents=summary_prompt,
        config=(
            genai_types.GenerateContentConfig(temperature=0.2)
            if genai_types is not None
            else None
        ),
    )
    return response.text


# ----------------------------------------------------------------------
# 지식베이스 확정 저장 완료 배너
# ----------------------------------------------------------------------
# render_confirm_to_kb_workspace()에서 저장 성공 직후 st.rerun()을 걸기
# 때문에, 이번 스크립트 실행에서는 그 성공 메시지가 사라진 상태로 시작함.
# 저장 경로를 세션에 잠깐 담아뒀던 값이 있으면 화면 맨 위에서 한 번 보여주고
# 곧바로 비워서, 다음 리런부터는 다시 뜨지 않게 함(한 번만 보이는 배너).
if st.session_state.get("_kb_last_saved_path"):
    st.success(f"지식베이스에 확정 저장되었습니다: {st.session_state['_kb_last_saved_path']}")
    st.session_state["_kb_last_saved_path"] = None


# ----------------------------------------------------------------------
# 새 주제 시작 / 현재 묶음 표시줄
# ----------------------------------------------------------------------
top_col1, top_col2 = st.columns([3, 1])
with top_col1:
    if st.session_state.current_thread:
        st.markdown(
            f"<span class='thread-badge'>진행 중인 대화: {get_thread_label(st.session_state.current_thread)} "
            f"({len(st.session_state.current_thread)}건)</span>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown("<span class='thread-badge'>새 대화 시작 전</span>", unsafe_allow_html=True)
with top_col2:
    if st.button("새 주제 시작", use_container_width=True, disabled=not st.session_state.current_thread):
        st.session_state.backlog.insert(0, {
            "started_at": st.session_state.current_thread[0]["time"],
            "turns": st.session_state.current_thread,
        })
        st.session_state.current_thread = []
        st.session_state.summary_doc = None
        st.rerun()


# ----------------------------------------------------------------------
# 질문 입력
# ----------------------------------------------------------------------
# 디자인 포인트(좌측 골드 바, 라벨 볼드 등)를 입히려면 CSS가 걸릴 클래스가
# 필요한데, st.form()의 위치 인자("question_form")는 st-key- 클래스를
# 만들어주는 key= 파라미터가 아니라서 CSS가 전혀 먹지 않는 문제가 있었음
# (실제 배포 화면에서 개발자도구로 textarea의 class 목록을 확인해보니
# st-key-question_form 클래스가 없었음 — 추측이 아니라 직접 확인된 사실).
# 로그인 폼(app_login_form)에 디자인이 정상 적용됐던 건 st.form 자체가
# 아니라, 그걸 감싼 st.container(key="pf_login_wrap") 덕분이었음.
# 그래서 동일하게 st.container(key=...)로 한 번 감싸는 방식으로 변경.
question_wrap = st.container(key="question_form_wrap")
with question_wrap:
    with st.form("question_form", clear_on_submit=True):
        placeholder_text = (
            "꼬리질문을 입력하세요 (이전 대화를 참고하여 답변합니다)"
            if st.session_state.current_thread
            else "예: 약국에서 의약품을 도매가로 구입해서 소매가로 판매할 때 부가세 처리는?"
        )
        user_question = st.text_area("세무 질의를 입력하세요", height=100, placeholder=placeholder_text)
        submitted = st.form_submit_button("조회", use_container_width=True)

if submitted:
    if not user_question.strip():
        st.warning("질문을 입력해주세요.")
    else:
        # 설계 의도 (2026-06-25 추가):
        # clear_on_submit=True라서 제출 즉시 입력칸이 비워지고 회색 예시 문구가
        # 다시 보이게 됨. 그 상태에서 "AI가 답변을 생성 중입니다..." 스피너만 돌면,
        # 방금 입력한 질문이 화면 어디에도 안 보여서 "내가 제대로 입력한 게
        # 맞나?"하는 불안감을 줄 수 있음. 그래서 로딩 중에도 방금 입력했던 질문을
        # 그대로 화면에 보여줘서, 무엇을 처리 중인지 명확하게 함.
        st.info(f"질문을 처리 중입니다:\n\n**{user_question.strip()}**")
        with st.spinner("AI가 답변을 생성 중입니다..."):
            answer = engine.generate_guideline_with_retry(
                user_question.strip(),
                thread_history=st.session_state.current_thread,
            )
        st.session_state.current_thread.append({
            "question": user_question.strip(),
            "answer": answer,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        st.session_state.summary_doc = None

        # 구글 시트 로깅 (설정된 경우에만 동작, 실패해도 화면 흐름에 영향 없음)
        if engine.sheet_logger and engine.sheet_logger.enabled:
            log_ts = engine.sheet_logger.log(
                question=user_question.strip(),
                answer=answer,
                user_type=user_type,
            )
            # 지식베이스 확정 시 검색기록 탭의 해당 행을 정확히 찾아 "확정됨"
            # 표시를 남기기 위한 값. log()가 실제로 사용한 타임스탬프를 그대로
            # 받아두는 게 안전함(위에서 별도로 만든 "time" 값과는 미세하게
            # (초 단위) 어긋날 수 있어 신뢰할 수 없음).
            st.session_state.current_thread[-1]["log_ts"] = log_ts

        st.rerun()


# ----------------------------------------------------------------------
# 현재 묶음 표시 (최신이 위)
# ----------------------------------------------------------------------
if st.session_state.current_thread:
    st.divider()
    st.subheader(f"현재 대화 ({len(st.session_state.current_thread)}건)")

    for idx, qa in enumerate(reversed(st.session_state.current_thread)):
        real_idx = len(st.session_state.current_thread) - 1 - idx
        # 설계 의도 (2026-06-25 수정):
        # 기존에는 real_idx(배열 위치)를 위젯 key로 썼는데, 이 값은 질문이
        # 추가되거나 "기록에서 제거"로 삭제되면 다른 항목들의 real_idx가 통째로
        # 바뀜. reversed()로 최신순 표시 중이라 화면 위치와 실제 인덱스가 계속
        # 달라지면서, session_state에 저장된 "이 카드의 확정 폼이 열려있다" 같은
        # 상태가 엉뚱한 카드로 옮겨붙는 문제가 있었음(꼬리질문을 2번 이상 하면
        # 펼쳐둔 폼이 갑자기 닫히는 등의 증상으로 나타날 수 있음).
        # 해결: time(타임스탬프, 초 단위)은 질문이 추가/삭제되어도 절대 바뀌지
        # 않는 고유값이므로, 이를 키의 기준으로 사용함.
        qa_key = qa["time"].replace("-", "").replace(":", "").replace(" ", "_")

        with st.container(border=True):
            st.markdown(f"<div class='qa-meta'>{qa['time']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='qa-question'>Q. {qa['question']}</div>", unsafe_allow_html=True)
            if st.session_state.get(f"cur_{qa_key}_kb_confirmed"):
                st.caption("✅ 지식베이스에 확정된 질문/답변입니다.")
            render_copyable_text(qa["answer"], key=f"copy_cur_{qa_key}")

            col1, col2, col3 = st.columns([1, 1, 1.4])
            with col1:
                save_show_key = f"save_show_{qa_key}"
                if st.button("이 답변 저장", key=f"save_cur_{qa_key}"):
                    st.session_state[save_show_key] = True
                if st.session_state.get(save_show_key):
                    offer_downloads(
                        title="세무질의 자문 기록",
                        question=qa["question"],
                        answer=qa["answer"],
                        filename_hint=f"자문_{qa_key}",
                        key_prefix=qa_key,
                    )
            with col2:
                if st.button("기록에서 제거", key=f"remove_cur_{qa_key}"):
                    st.session_state.current_thread.pop(real_idx)
                    st.rerun()
            with col3:
                if is_admin:
                    render_confirm_to_kb_button(
                        question=qa["question"], answer=qa["answer"], key_prefix=f"cur_{qa_key}",
                        source_log_timestamp=qa.get("log_ts"),
                        already_confirmed=bool(st.session_state.get(f"cur_{qa_key}_kb_confirmed")),
                    )

    # ------------------------------------------------------------------
    # 전체 종합 (현재 묶음)
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("현재 대화 종합")
    st.caption("지금 진행 중인 대화 묶음만 AI가 하나의 완성된 자문 문서로 정리합니다.")

    if st.button("현재 대화 종합 문서 생성", type="primary", use_container_width=True):
        with st.spinner("지금까지의 대화를 종합하는 중입니다..."):
            st.session_state.summary_doc = run_summary(st.session_state.current_thread)
            st.session_state.summary_doc_turns_count = len(st.session_state.current_thread)
            # 지식베이스 확정 저장 시 "질문" 자리에 대신 표시할 라벨.
            # 종합 문서는 여러 질문을 묶은 것이라 단일 질문이 없으므로,
            # 첫 질문을 짧게 잘라 대표 제목처럼 사용함(get_thread_label과 동일 방식).
            st.session_state.summary_doc_label = (
                f"[종합문서] {get_thread_label(st.session_state.current_thread)} "
                f"외 {st.session_state.summary_doc_turns_count}건 종합"
            )
        # 설계 의도 (2026-06-25 변경): 종합 문서는 '저장' 버튼을 누를 때가 아니라
        # '생성'하는 시점에 자동으로 구글 시트(종합문서 탭)에 기록함. 개별 질문이
        # 답변 생성 즉시 자동 기록되는 것과 동일한 방식으로 통일하기 위함.
        # (이전에는 "종합 문서 저장" 버튼을 직접 눌렀을 때만 기록했었음)
        if engine.sheet_logger and engine.sheet_logger.enabled:
            st.session_state.summary_doc_source_ts = engine.sheet_logger.log_summary(
                turns_count=st.session_state.summary_doc_turns_count,
                summary_text=st.session_state.summary_doc,
                user_type=user_type,
            )
        else:
            st.session_state.summary_doc_source_ts = None
        # 새로 생성한 문서이므로 이전 확정 배지는 초기화(아직 확정 전 상태)
        st.session_state.pop("summary_cur_kb_confirmed", None)

    if st.session_state.summary_doc:
        st.markdown("### 종합 문서 결과")
        if st.session_state.get("summary_cur_kb_confirmed"):
            st.success("✅ 지식베이스에 확정된 문서입니다.")
        else:
            st.warning("⚠️ 아직 지식베이스에 확정되지 않은 AI 생성 문서입니다. 검증 절차를 거치지 않아 오류가 있을 수 있습니다.")
        render_copyable_text(st.session_state.summary_doc, key="copy_summary_cur")
        save_col, kb_col = st.columns([1, 1.4])
        with save_col:
            save_show_key = "save_summary_cur_show"
            if st.button("종합 문서 저장", key="save_summary_cur"):
                st.session_state[save_show_key] = True
            if st.session_state.get(save_show_key):
                offer_downloads(
                    title="세무질의 종합 자문 문서",
                    question="",
                    answer=st.session_state.summary_doc,
                    filename_hint=f"종합자문_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
                    key_prefix="save_summary_cur",
                )
        with kb_col:
            # 설계 의도 (종합문서 지식베이스 확정 기능, 신규):
            # 개별 답변 각각을 확정 저장하기보다, 꼬리질문을 거치며 정정된
            # 최종 결론이 반영된 종합 문서 하나를 확정하는 편이 지식베이스
            # 품질에 더 유리하다는 판단에 따라 추가함(개별 확정 버튼과 동일한
            # render_confirm_to_kb_button/workspace를 그대로 재사용 —
            # question/answer가 순수 텍스트이기만 하면 검증·수정·저장 흐름은
            # 개별 답변이든 종합 문서든 동일하게 동작함).
            if is_admin:
                render_confirm_to_kb_button(
                    question=st.session_state.get("summary_doc_label", "종합 문서"),
                    answer=st.session_state.summary_doc,
                    key_prefix="summary_cur",
                    source_summary_timestamp=st.session_state.get("summary_doc_source_ts"),
                    already_confirmed=bool(st.session_state.get("summary_cur_kb_confirmed")),
                )
else:
    st.info("새 질문을 입력하면 여기서 대화가 시작됩니다.")


# ----------------------------------------------------------------------
# 백로그 (이전 대화 묶음들)
# ----------------------------------------------------------------------
if st.session_state.backlog:
    st.divider()
    with st.expander(f"이전 대화 보기 ({len(st.session_state.backlog)}개 묶음)", expanded=False):
        for b_idx, bundle in enumerate(st.session_state.backlog):
            label = f"{bundle['started_at']} — {get_thread_label(bundle['turns'])} ({len(bundle['turns'])}건)"
            with st.expander(label, expanded=False):
                for qa in bundle["turns"]:
                    st.markdown(f"<div class='qa-meta'>{qa['time']}</div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='qa-question'>Q. {qa['question']}</div>", unsafe_allow_html=True)
                    qa_time_key = qa["time"].replace("-", "").replace(":", "").replace(" ", "_")
                    render_copyable_text(qa["answer"], key=f"copy_backlog_{b_idx}_{qa_time_key}")
                    st.markdown("---")

                col1, col2 = st.columns([1, 1])
                with col1:
                    backlog_summary_key = f"backlog_summary_{b_idx}"
                    if st.button("이 묶음 종합 문서 생성", key=f"summary_backlog_{b_idx}"):
                        with st.spinner("종합하는 중입니다..."):
                            summary_text = run_summary(bundle["turns"])
                        # 결과를 session_state에 저장해야 다음 리런(예: 다운로드
                        # 버튼 클릭)에서도 화면에 계속 남아있음. 이전에는 지역
                        # 변수에만 담아뒀어서, do_save() 안에서 예외가 나거나
                        # 다른 위젯이 또 눌리면 이 결과 자체가 사라져 버렸음.
                        st.session_state[backlog_summary_key] = summary_text
                        # 종합 문서는 생성 시점에 자동으로 구글 시트(종합문서 탭)에 기록.
                        if engine.sheet_logger and engine.sheet_logger.enabled:
                            st.session_state[f"{backlog_summary_key}_source_ts"] = engine.sheet_logger.log_summary(
                                turns_count=len(bundle["turns"]),
                                summary_text=summary_text,
                                user_type=user_type,
                            )
                        else:
                            st.session_state[f"{backlog_summary_key}_source_ts"] = None
                        st.session_state.pop(f"{backlog_summary_key}_kb_confirmed", None)

                    if st.session_state.get(backlog_summary_key):
                        if st.session_state.get(f"{backlog_summary_key}_kb_confirmed"):
                            st.success("✅ 지식베이스에 확정된 문서입니다.")
                        else:
                            st.warning("⚠️ 아직 지식베이스에 확정되지 않은 AI 생성 문서입니다. 검증 절차를 거치지 않아 오류가 있을 수 있습니다.")
                        render_copyable_text(st.session_state[backlog_summary_key], key=f"{backlog_summary_key}_copy_text")
                        save_show_key = f"{backlog_summary_key}_save_show"
                        if st.button("이 종합 문서 저장", key=f"{backlog_summary_key}_save_btn"):
                            st.session_state[save_show_key] = True
                        if st.session_state.get(save_show_key):
                            offer_downloads(
                                title="세무질의 종합 자문 문서",
                                question="",
                                answer=st.session_state[backlog_summary_key],
                                filename_hint=f"종합자문_{bundle['started_at'].replace('-', '').replace(':', '').replace(' ', '_')}",
                                key_prefix=backlog_summary_key,
                            )
                        # 이전 대화 묶음의 종합 문서도 현재 대화 종합과 동일하게
                        # 지식베이스 확정 저장 가능하도록 함.
                        if is_admin:
                            render_confirm_to_kb_button(
                                question=(
                                    f"[종합문서] {get_thread_label(bundle['turns'])} "
                                    f"외 {len(bundle['turns'])}건 종합"
                                ),
                                answer=st.session_state[backlog_summary_key],
                                key_prefix=backlog_summary_key,
                                source_summary_timestamp=st.session_state.get(f"{backlog_summary_key}_source_ts"),
                                already_confirmed=bool(st.session_state.get(f"{backlog_summary_key}_kb_confirmed")),
                            )
                with col2:
                    if st.button("이 묶음 삭제", key=f"delete_backlog_{b_idx}"):
                        st.session_state.backlog.pop(b_idx)
                        st.rerun()

    # 전체(백로그+현재) 종합
    st.subheader("전체 종합 (현재 + 이전 대화 모두)")
    if st.button("전체 묶음 종합 문서 생성", use_container_width=True):
        with st.spinner("모든 대화를 종합하는 중입니다..."):
            all_turns = []
            for bundle in reversed(st.session_state.backlog):  # 오래된 순
                all_turns.extend(bundle["turns"])
            all_turns.extend(st.session_state.current_thread)
            st.session_state.summary_doc = run_summary(all_turns)
            st.session_state.summary_doc_turns_count = len(all_turns)
            st.session_state.summary_doc_label = (
                f"[종합문서] {get_thread_label(all_turns)} "
                f"외 {st.session_state.summary_doc_turns_count}건 종합 (전체)"
            )
        # "현재 대화 종합"과 동일하게, 생성 시점에 자동으로 구글 시트에 기록.
        if engine.sheet_logger and engine.sheet_logger.enabled:
            st.session_state.summary_doc_source_ts = engine.sheet_logger.log_summary(
                turns_count=st.session_state.summary_doc_turns_count,
                summary_text=st.session_state.summary_doc,
                user_type=user_type,
            )
        else:
            st.session_state.summary_doc_source_ts = None
        st.session_state.pop("summary_cur_kb_confirmed", None)
        st.rerun()


# ----------------------------------------------------------------------
# 지식베이스 확정 저장 작업 공간 (화면 맨 아래, 화면 전체 너비)
# ----------------------------------------------------------------------
# 어디서든(검색 기록, 현재 대화 카드 등) "지식베이스에 확정 저장" 버튼을 누르면
# 여기에 작업 대상이 지정되고, 이 섹션에 검증/수정/저장 흐름이 펼쳐짐.
# is_admin으로 한 번 더 감싸는 이유: 위에서 버튼 자체를 직원에게는 안 보이게
# 했지만(렌더링되는 곳들에서 if is_admin으로 막음), 혹시라도 세션에
# kb_confirm_target이 남아있는 상태로 직원 계정으로 로그인하는 경우까지
# 대비해 이 작업 공간 자체도 관리자가 아니면 절대 그려지지 않도록 이중으로 막음.
if is_admin:
    render_confirm_to_kb_workspace()
