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

실행 방법:
    streamlit run streamlit_ui.py

Windows 환경 기준으로 작성됨.
"""

import os
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
       포스원 세무 자문 AI 시스템 - 디자인 테마
       기존 위젯 동작(버튼, 폼, 인풋)은 그대로 두고 색/모양만 입힌 CSS.
       Streamlit 기본 마크업 구조에 의존하므로, 버전 업그레이드 시
       data-testid 셀렉터가 바뀌면 일부 효과가 사라질 수 있음(동작에는
       영향 없음 — 순수 스타일 레이어).
       ------------------------------------------------------------------ */

    :root {
        --pf-accent: #2563eb;
        --pf-accent-dark: #1d4ed8;
        --pf-accent-bg: #eff6ff;
        --pf-accent-border: #bfdbfe;
        --pf-text-strong: #1e3a5f;
        --pf-text-muted: #6b7280;
        --pf-border: #e5e7eb;
        --pf-success-bg: #ecfdf5;
        --pf-success-border: #a7f3d0;
        --pf-success-text: #047857;
    }

    .block-container { max-width: 880px; padding-top: 2.2rem; }

    /* 페이지 타이틀 영역 ------------------------------------------------ */
    .pf-header-row {
        display: flex; align-items: center; gap: 12px; margin-bottom: 0.2rem;
    }
    .pf-header-icon {
        width: 38px; height: 38px; border-radius: 10px;
        background: var(--pf-accent-bg); border: 1px solid var(--pf-accent-border);
        display: flex; align-items: center; justify-content: center;
        font-size: 19px; flex-shrink: 0;
    }
    .pf-header-title {
        font-size: 1.55rem; font-weight: 700; color: var(--pf-text-strong);
        margin: 0; line-height: 1.25;
    }
    .pf-header-caption {
        color: var(--pf-text-muted); font-size: 0.92rem; margin: 0.1rem 0 0 50px;
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

    /* 진행 중인 대화 배지 ------------------------------------------------- */
    .thread-badge {
        display: inline-flex; align-items: center; gap: 6px;
        background: var(--pf-accent-bg); color: var(--pf-accent-dark);
        border: 1px solid var(--pf-accent-border);
        border-radius: 999px; padding: 5px 14px; font-size: 0.82rem;
        font-weight: 600;
    }

    /* st.container(border=True) 카드 -> 살짝 더 또렷하게 */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        border-color: var(--pf-border) !important;
    }

    /* primary 버튼(조회, 확정 저장 실행 등) -> 포인트 컬러로 채움 */
    button[kind="primary"], button[kind="primaryFormSubmit"] {
        background-color: var(--pf-accent) !important;
        border-color: var(--pf-accent) !important;
        color: #fff !important;
    }
    button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
        background-color: var(--pf-accent-dark) !important;
        border-color: var(--pf-accent-dark) !important;
    }

    /* st.info / st.success / st.warning 박스 라운드 처리 */
    div[data-testid="stAlert"] {
        border-radius: 10px !important;
    }

    /* 사이드바 섹션 제목과의 간격 정리 */
    section[data-testid="stSidebar"] h3 {
        margin-top: 0.3rem;
    }

    /* 구분선 여백 약간 축소 (섹션이 많아 답답해 보이는 것 방지) */
    hr { margin: 1.1rem 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="pf-header-row">
        <div class="pf-header-icon">🧮</div>
        <p class="pf-header-title">포스원 세무 자문 AI 시스템</p>
    </div>
    <p class="pf-header-caption">기장 직원 / 회계사를 위한 세무질의 실시간 응답 도구</p>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# 로그인 (사내 공통 비밀번호)
# ----------------------------------------------------------------------
# 설계 의도 (2026-06-25 추가):
# - 이 화면을 웹(Streamlit Community Cloud 등)에 올리면 주소만 알면 누구나 접근 가능해짐.
#   그래서 "사내 직원만 들어올 수 있는 문" 역할로 공통 비밀번호 1개를 둠.
# - 이건 PIN(사이드바의 '확정 PIN', 회계사가 지식베이스에 확정 저장할 때 쓰는 것)과는
#   완전히 다른 레이어임. 여기 비밀번호는 "웹사이트 입장 가능 여부"만 가르고,
#   PIN은 "지식베이스 확정 저장 가능 여부"를 가름.
# - 비밀번호 값은 코드에 직접 적지 않고 .env(로컬) 또는 Streamlit Secrets(웹 배포 시)의
#   APP_PASSWORD 값으로 관리함.
# - 추후 "직원용/관리자용 비밀번호를 따로 분리"하는 더 세밀한 권한 구조로 발전시킬 수 있으나,
#   현재는 가장 단순한 형태(공통 비밀번호 1개)로만 구현함.
def check_app_password() -> bool:
    """입력한 비밀번호가 APP_PASSWORD와 일치하면 True. 세션 동안 결과를 유지."""
    if st.session_state.get("app_authenticated"):
        return True

    correct_password = os.getenv("APP_PASSWORD", "")

    if not correct_password:
        # APP_PASSWORD가 아예 설정되지 않은 경우: 로컬 테스트 등에서 막히지 않도록
        # 안내만 띄우고 통과시킴. 단, 웹에 배포할 때는 반드시 설정해야 함.
        st.warning(
            "APP_PASSWORD가 설정되어 있지 않습니다. 로컬 환경에서는 그대로 진행되지만, "
            "웹에 배포할 때는 .env(로컬) 또는 Streamlit Secrets(웹)에 APP_PASSWORD를 "
            "반드시 설정해야 합니다."
        )
        return True

    st.subheader("사내 로그인")
    st.caption("포스원 회계법인 직원 전용 화면입니다. 사내 공통 비밀번호를 입력해주세요.")

    with st.form("app_login_form"):
        pw_input = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("입장")

    if submitted:
        if pw_input == correct_password:
            st.session_state.app_authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")

    return False


if not check_app_password():
    st.stop()  # 비밀번호가 맞을 때까지 아래 모든 코드(엔진 초기화, 화면 등)를 실행하지 않음


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
st.session_state.setdefault("kb_confirm_target", None)


def get_thread_label(turns):
    """묶음의 첫 질문을 짧게 잘라 제목처럼 사용"""
    if not turns:
        return "(빈 대화)"
    first_q = turns[0]["question"]
    return first_q[:40] + ("..." if len(first_q) > 40 else "")


def render_confirm_to_kb_button(question: str, answer: str, key_prefix: str, dialog_row_key: str = None):
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
    """
    target_key = "kb_confirm_target"
    if st.button("지식베이스에 확정 저장", key=f"{key_prefix}_confirm_entry_btn"):
        st.session_state[target_key] = {
            "question": question,
            "answer": answer,
            "key_prefix": key_prefix,
        }
        # 이전에 다른 항목을 검증/수정하던 상태가 남아있으면 깨끗하게 초기화
        st.session_state[f"{key_prefix}_verification_result"] = None
        st.session_state[f"{key_prefix}_edited_content"] = answer

        if dialog_row_key:
            # 다이얼로그 안에서 호출된 경우: 다이얼로그를 닫고 메인 화면으로 이동
            st.session_state[dialog_row_key] = None
            st.rerun()
        else:
            st.info("화면 맨 아래 '지식베이스 확정 저장 작업 공간'으로 이동해 진행해주세요.")


def render_confirm_to_kb_workspace():
    """
    실제 검증 → 수정 → 저장 흐름이 펼쳐지는 전용 작업 공간.
    render_confirm_to_kb_button으로 대상이 지정된 경우에만 화면에 나타나며,
    메인 화면 맨 아래(화면 전체 너비)에 한 번만 그려짐.
    """
    target = st.session_state.get("kb_confirm_target")
    if not target:
        return

    question = target["question"]
    answer = target["answer"]
    key_prefix = target["key_prefix"]

    st.divider()
    st.header("지식베이스 확정 저장 작업 공간")
    st.caption("아래에서 검증 결과를 확인하고, 필요한 부분을 직접 수정한 뒤 PIN으로 최종 승인하세요.")

    with st.expander("대상 질문 / 원본 답변 보기", expanded=False):
        st.markdown(f"**질문**: {question}")
        st.markdown("**원본 답변**")
        st.markdown(answer)

    if not engine.has_pin_set():
        st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
        return

    verified_key = f"{key_prefix}_verification_result"
    edited_key = f"{key_prefix}_edited_content"

    # 1단계: 자동 검증
    if st.session_state.get(verified_key) is None:
        if st.button("① 내용 검증하기 (웹검색)", key=f"{key_prefix}_run_verify_btn", type="primary"):
            with st.spinner("답변 내용을 웹검색으로 재검증하는 중입니다..."):
                st.session_state[verified_key] = engine.verify_before_confirm(question, answer)
            st.rerun()
        else:
            st.caption("저장 전에 먼저 내용을 검증해주세요. (검증 결과는 참고용이며, 최종 판단은 회계사가 직접 합니다.)")
            return

    verification = st.session_state[verified_key]

    st.markdown("### ② 검증 및 자동 수정 결과")
    st.markdown(f"**수정 사항**: {verification['correction_summary']}")
    with st.expander("검증 상세 내용 보기 (검색 근거 등)", expanded=False):
        st.info(verification["verification_text"])
    st.caption(f"추천 저장 파일: {verification['recommended_file']} — {verification['recommended_reason']}")

    # 2단계: 저장할 내용 확인/수정
    # 설계 의도 (2026-06-25 추가 — 검증 결과 자동 반영):
    # 기본값을 원본(answer)이 아니라 AI가 검증 후 직접 고친 최종본
    # (verification["corrected_content"])으로 설정함. 회계사가 매번 검증 결과를
    # 읽고 원문에서 같은 부분을 직접 찾아 고치는 수고를 줄이기 위함. 다만 자동
    # 수정이 100% 정확하다고 보장할 수 없으므로, 텍스트칸은 여전히 자유롭게
    # 추가로 손볼 수 있게 열어둠.
    if edited_key not in st.session_state:
        st.session_state[edited_key] = verification["corrected_content"]

    st.markdown("### ③ 저장할 내용 확인/수정")
    st.caption("위 '수정 사항'이 이미 아래 내용에 반영되어 있습니다. 한번 훑어보고 필요하면 추가로 고쳐주세요.")
    st.session_state[edited_key] = st.text_area(
        "지식베이스에 저장될 최종 내용",
        value=st.session_state[edited_key],
        height=500,
        key=f"{key_prefix}_edit_textarea",
    )

    # 3단계: 저장 파일 선택
    default_idx = (
        engine.KNOWLEDGE_FILE_OPTIONS.index(verification["recommended_file"])
        if verification["recommended_file"] in engine.KNOWLEDGE_FILE_OPTIONS
        else 0
    )
    target_file = st.selectbox(
        "④ 저장할 지식베이스 파일",
        options=engine.KNOWLEDGE_FILE_OPTIONS,
        index=default_idx,
        key=f"{key_prefix}_target_file_select",
    )

    # 4단계: PIN 입력 후 최종 저장
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
                st.success(f"지식베이스에 확정 저장되었습니다: {saved_path}")
                st.session_state["kb_confirm_target"] = None
                st.session_state[verified_key] = None
            else:
                st.error("PIN이 일치하지 않습니다.")
        if cancelled:
            st.session_state["kb_confirm_target"] = None
            st.session_state[verified_key] = None
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
    st.divider()
    st.markdown(answer)
    st.divider()

    dialog_key_base = f"dialog_{row.get('일시', '')}".replace(" ", "_").replace(":", "")

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
        render_confirm_to_kb_button(
            question=question, answer=answer, key_prefix=dialog_key_base,
            dialog_row_key="_dialog_log_row",
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
    (종합 문서는 여러 질문을 재구성한 2차 가공물이므로 지식베이스 확정 대상에서는 제외)
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
    st.markdown(summary_text)
    st.divider()

    dialog_key_base = f"sdialog_{row.get('일시', '')}".replace(" ", "_").replace(":", "")

    col1, col2 = st.columns([1, 1])
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
    st.subheader("사용자 구분")
    user_type = st.radio(
        "현재 사용자",
        options=["직원", "회계사"],
        horizontal=True,
        help="검색 기록에 누가 질문했는지 표시하기 위한 구분입니다 (로그인 기능은 아닙니다)",
    )

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

    if st.button(
        "지식 베이스 새로고침",
        help="로컬 _knowledge 폴더 파일을 수정했거나, 다른 곳에서 확정 저장한 "
        "구글시트 지식베이스 내용을 즉시 반영하고 싶을 때 누르세요.",
    ):
        engine.load_knowledge_base(force_reload=True)
        st.success("지식 베이스를 다시 불러왔습니다.")

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
                label = f"{row.get('일시', '')} | {row.get('사용자구분', '')} | {row.get('질문', '')[:30]}"
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
        "'종합 문서 저장' 버튼을 직접 눌렀을 때만 여기에 남습니다 "
        "(개별 질문 기록과는 별도 탭에 저장되어 섞이지 않습니다)."
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
                label = (
                    f"{row.get('일시', '')} | 질의 {row.get('포함된질의건수', '?')}건 | "
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
        st.code(
            f"[구글 시트 관련]\n"
            f"GOOGLE_SHEET_ID 읽힘: {'예 (' + str(len(_gsheet_id)) + '자)' if _gsheet_id else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_JSON 읽힘: {'예 (' + str(len(_gcred_json)) + '자)' if _gcred_json else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_PATH 읽힘: {'예 (' + str(len(_gcred_path)) + '자)' if _gcred_path else '아니오 (빈 값)'}\n"
            f"sheet_logger.enabled: {engine.sheet_logger.enabled if engine.sheet_logger else '(sheet_logger 없음)'}\n"
            f"sheet_logger.error_message: {_err}\n"
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
    save_formats = st.multiselect(
        "다운로드 형식 (여러 개 선택 가능)",
        options=["md", "docx", "pdf"],
        default=["md"],
        help="pdf는 서버에 한글 폰트가 없으면 생성되지 않을 수 있습니다. "
        "이 경우 md 또는 docx로 받아주세요.",
    )
    use_custom_filename = st.checkbox("파일명 직접 지정", value=False)
    custom_filename = ""
    if use_custom_filename:
        custom_filename = st.text_input("파일명 (확장자 제외)", value="")

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
    filename = safe_filename(
        custom_filename.strip() if (use_custom_filename and custom_filename.strip()) else filename_hint
    )
    formats = save_formats or ["md"]

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
        with st.spinner("AI가 답변을 생성 중입니다 (3~5초 소요)..."):
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
            engine.sheet_logger.log(
                question=user_question.strip(),
                answer=answer,
                user_type=user_type,
            )

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
            st.markdown(qa["answer"])

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
                render_confirm_to_kb_button(
                    question=qa["question"], answer=qa["answer"], key_prefix=f"cur_{qa_key}"
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
        # 설계 의도 (2026-06-25 변경): 종합 문서는 '저장' 버튼을 누를 때가 아니라
        # '생성'하는 시점에 자동으로 구글 시트(종합문서 탭)에 기록함. 개별 질문이
        # 답변 생성 즉시 자동 기록되는 것과 동일한 방식으로 통일하기 위함.
        # (이전에는 "종합 문서 저장" 버튼을 직접 눌렀을 때만 기록했었음)
        if engine.sheet_logger and engine.sheet_logger.enabled:
            engine.sheet_logger.log_summary(
                turns_count=st.session_state.summary_doc_turns_count,
                summary_text=st.session_state.summary_doc,
                user_type=user_type,
            )

    if st.session_state.summary_doc:
        st.markdown("### 종합 문서 결과")
        st.markdown(st.session_state.summary_doc)
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
                    st.markdown(qa["answer"])
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
                            engine.sheet_logger.log_summary(
                                turns_count=len(bundle["turns"]),
                                summary_text=summary_text,
                                user_type=user_type,
                            )

                    if st.session_state.get(backlog_summary_key):
                        st.markdown(st.session_state[backlog_summary_key])
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
        # "현재 대화 종합"과 동일하게, 생성 시점에 자동으로 구글 시트에 기록.
        if engine.sheet_logger and engine.sheet_logger.enabled:
            engine.sheet_logger.log_summary(
                turns_count=st.session_state.summary_doc_turns_count,
                summary_text=st.session_state.summary_doc,
                user_type=user_type,
            )
        st.rerun()


# ----------------------------------------------------------------------
# 지식베이스 확정 저장 작업 공간 (화면 맨 아래, 화면 전체 너비)
# ----------------------------------------------------------------------
# 어디서든(검색 기록, 현재 대화 카드 등) "지식베이스에 확정 저장" 버튼을 누르면
# 여기에 작업 대상이 지정되고, 이 섹션에 검증/수정/저장 흐름이 펼쳐짐.
render_confirm_to_kb_workspace()
