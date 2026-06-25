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
from doc_export import export_response


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
    .block-container { max-width: 880px; }
    .qa-question { font-weight: 600; color: #1f3a5f; margin-bottom: 0.2rem; }
    .qa-meta { color: #888; font-size: 0.8rem; }
    .thread-badge {
        display: inline-block; background: #eef3fa; color: #1f3a5f;
        border-radius: 4px; padding: 2px 8px; font-size: 0.78rem; margin-bottom: 6px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("포스원 세무 자문 AI 시스템")
st.caption("기장 직원/회계사를 위한 세무질의 실시간 응답 도구")


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
st.session_state.setdefault(
    "output_dir_value",
    os.getenv("OUTPUT_DIR", "").strip() or str(Path(__file__).resolve().parent / "세법검토_아카이브"),
)


def get_thread_label(turns):
    """묶음의 첫 질문을 짧게 잘라 제목처럼 사용"""
    if not turns:
        return "(빈 대화)"
    first_q = turns[0]["question"]
    return first_q[:40] + ("..." if len(first_q) > 40 else "")


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

    col1, col2 = st.columns([1, 1.4])
    with col1:
        if st.button("이 답변 저장", key=f"{dialog_key_base}_save"):
            do_save(
                title="세무질의 자문 기록",
                question=question,
                answer=answer,
                filename_hint=f"자문_{row.get('일시', '').replace('-', '').replace(':', '').replace(' ', '_')}",
            )

    with col2:
        confirm_key = f"{dialog_key_base}_show_confirm"
        if st.button("지식베이스에 확정 저장", key=f"{dialog_key_base}_confirm_btn"):
            st.session_state[confirm_key] = True

        if st.session_state.get(confirm_key):
            if not engine.has_pin_set():
                st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
            else:
                with st.form(f"{dialog_key_base}_confirm_form"):
                    pin_input = st.text_input(
                        "회계사 확정 PIN", type="password", key=f"{dialog_key_base}_pin_input"
                    )
                    target_file = st.selectbox(
                        "추가할 지식베이스 파일",
                        options=[
                            "01_공통_세무질의회신집.txt",
                            "02_업종별_기장유의사항.txt",
                            "04_부가세_처리지침.txt",
                            "05_기타_세법_예규.txt",
                        ],
                        key=f"{dialog_key_base}_target_file",
                    )
                    if st.form_submit_button("확정 저장 실행"):
                        if engine.verify_pin(pin_input):
                            saved_path = engine.confirm_to_knowledge_base(
                                question=question,
                                confirmed_content=answer,
                                target_file=target_file,
                            )
                            st.success(f"지식베이스에 확정 저장되었습니다: {saved_path}")
                            st.session_state[confirm_key] = False
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
    if st.button("이 종합 문서 다시 저장", key=f"{dialog_key_base}_save"):
        do_save(
            title="세무질의 종합 자문 문서",
            question="",
            answer=summary_text,
            filename_hint=f"종합자문_{row.get('일시', '').replace('-', '').replace(':', '').replace(' ', '_')}",
        )


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

    if st.button("지식 베이스 새로고침", help="_knowledge 폴더 파일을 수정한 뒤 누르면 즉시 반영됩니다"):
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
        st.code(
            f"[구글 시트 관련]\n"
            f"GOOGLE_SHEET_ID 읽힘: {'예 (' + str(len(_gsheet_id)) + '자)' if _gsheet_id else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_JSON 읽힘: {'예 (' + str(len(_gcred_json)) + '자)' if _gcred_json else '아니오 (빈 값)'}\n"
            f"GOOGLE_CREDENTIALS_PATH 읽힘: {'예 (' + str(len(_gcred_path)) + '자)' if _gcred_path else '아니오 (빈 값)'}\n"
            f"sheet_logger.enabled: {engine.sheet_logger.enabled if engine.sheet_logger else '(sheet_logger 없음)'}\n"
            f"sheet_logger.error_message: {engine.sheet_logger.error_message if engine.sheet_logger else '(sheet_logger 없음)'!r}\n"
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
    output_dir = st.text_input("저장 폴더", key="output_dir_value")
    save_formats = st.multiselect(
        "저장 형식 (여러 개 선택 가능)",
        options=["md", "docx", "pdf"],
        default=["md"],
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
def do_save(title: str, question: str, answer: str, filename_hint: str = None, is_summary: bool = False, turns_count: int = None):
    """
    결과를 로컬 파일(md/docx/pdf)로 저장.

    is_summary=True인 경우(= '종합 문서 저장' 버튼을 직접 눌렀을 때만), 구글 시트의
    '종합문서' 탭에도 함께 기록함. 개별 질문 저장(is_summary=False)은 구글 시트에
    별도로 남기지 않음 — 개별 질문은 답변 생성 시점에 이미 자동으로 기록되어 있으므로
    저장 버튼을 누를 때 또 남기면 중복이 됨.
    """
    filename = custom_filename.strip() if (use_custom_filename and custom_filename.strip()) else filename_hint
    formats = save_formats or ["md"]

    result = export_response(
        title=title,
        question=question,
        response_md=answer,
        output_dir=output_dir,
        filename=filename,
        formats=formats,
    )

    saved_paths = [str(p) for fmt, p in result.items() if fmt != "errors" and p]
    if saved_paths:
        st.success("저장 완료:\n" + "\n".join(f"- {p}" for p in saved_paths))
    if result["errors"]:
        for err in result["errors"]:
            st.warning(err)

    if is_summary and engine.sheet_logger and engine.sheet_logger.enabled:
        engine.sheet_logger.log_summary(
            turns_count=turns_count if turns_count is not None else 0,
            summary_text=answer,
            user_type=user_type,
        )


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
        with st.container(border=True):
            st.markdown(f"<div class='qa-meta'>{qa['time']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='qa-question'>Q. {qa['question']}</div>", unsafe_allow_html=True)
            st.markdown(qa["answer"])

            col1, col2, col3 = st.columns([1, 1, 1.4])
            with col1:
                if st.button("이 답변 저장", key=f"save_cur_{real_idx}"):
                    do_save(
                        title="세무질의 자문 기록",
                        question=qa["question"],
                        answer=qa["answer"],
                        filename_hint=f"자문_{qa['time'].replace('-', '').replace(':', '').replace(' ', '_')}",
                    )
            with col2:
                if st.button("기록에서 제거", key=f"remove_cur_{real_idx}"):
                    st.session_state.current_thread.pop(real_idx)
                    st.rerun()
            with col3:
                confirm_key = f"show_confirm_{real_idx}"
                if st.button("지식베이스에 확정 저장", key=f"confirm_btn_{real_idx}"):
                    st.session_state[confirm_key] = True

                if st.session_state.get(confirm_key):
                    if not engine.has_pin_set():
                        st.warning("먼저 사이드바에서 확정 PIN을 설정해주세요.")
                    else:
                        with st.form(f"confirm_form_{real_idx}"):
                            pin_input = st.text_input("회계사 확정 PIN", type="password", key=f"pin_input_{real_idx}")
                            target_file = st.selectbox(
                                "추가할 지식베이스 파일",
                                options=[
                                    "01_공통_세무질의회신집.txt",
                                    "02_업종별_기장유의사항.txt",
                                    "04_부가세_처리지침.txt",
                                    "05_기타_세법_예규.txt",
                                ],
                                key=f"target_file_{real_idx}",
                            )
                            if st.form_submit_button("확정 저장 실행"):
                                if engine.verify_pin(pin_input):
                                    saved_path = engine.confirm_to_knowledge_base(
                                        question=qa["question"],
                                        confirmed_content=qa["answer"],
                                        target_file=target_file,
                                    )
                                    st.success(f"지식베이스에 확정 저장되었습니다: {saved_path}")
                                    st.session_state[confirm_key] = False
                                else:
                                    st.error("PIN이 일치하지 않습니다.")

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

    if st.session_state.summary_doc:
        st.markdown("### 종합 문서 결과")
        st.markdown(st.session_state.summary_doc)
        if st.button("종합 문서 저장", key="save_summary_cur"):
            do_save(
                title="세무질의 종합 자문 문서",
                question="",
                answer=st.session_state.summary_doc,
                filename_hint=f"종합자문_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
                is_summary=True,
                turns_count=st.session_state.summary_doc_turns_count,
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
                    if st.button("이 묶음 종합 문서 생성", key=f"summary_backlog_{b_idx}"):
                        with st.spinner("종합하는 중입니다..."):
                            summary_text = run_summary(bundle["turns"])
                        st.markdown(summary_text)
                        do_save(
                            title="세무질의 종합 자문 문서",
                            question="",
                            answer=summary_text,
                            filename_hint=f"종합자문_{bundle['started_at'].replace('-', '').replace(':', '').replace(' ', '_')}",
                            is_summary=True,
                            turns_count=len(bundle["turns"]),
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
        st.rerun()
