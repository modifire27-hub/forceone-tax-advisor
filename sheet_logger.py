# -*- coding: utf-8 -*-
"""
sheet_logger.py
================
세무 질의응답 기록을 구글 스프레드시트에 자동으로 누적 저장하는 모듈.

설계 원칙:
- 선택 기능: .env에 GOOGLE_SHEET_ID와 GOOGLE_CREDENTIALS_PATH가 모두 설정된 경우에만 작동.
  설정이 없으면 로깅 없이 기존처럼 동작 (하위 호환).
- 로깅 실패가 절대 답변 생성 자체를 막아서는 안 됨 — 모든 호출은 예외를 삼키고
  경고만 출력한 뒤 계속 진행.
- 한 행 = 한 건의 질의응답. 시간순으로 계속 누적됨 (덮어쓰지 않음).

사전 준비 (Windows 환경, 1회만):
1. https://console.cloud.google.com 에서 프로젝트 생성
2. "Google Sheets API"와 "Google Drive API" 활성화
3. "서비스 계정" 생성 → JSON 키 파일 다운로드
4. 구글 스프레드시트를 새로 만들고, 그 시트를 위 서비스 계정 이메일과 "편집자"로 공유
5. 시트 URL에서 ID 부분을 복사 (예: docs.google.com/spreadsheets/d/이부분/edit)
6. .env에 다음을 설정:
   GOOGLE_SHEET_ID=시트ID
   GOOGLE_CREDENTIALS_PATH=다운로드한_JSON_키_파일_경로

Windows 환경 기준으로 작성됨.
"""

import os
import hashlib
from pathlib import Path
from datetime import datetime


class SheetLogger:
    """구글 스프레드시트 로깅 클라이언트. 설정 미비 시 비활성 상태로 안전하게 동작."""

    HEADER = ["일시", "질문", "답변요약", "전체답변", "근거유형", "사용자구분"]

    # 종합문서 탭(워크시트) 이름 및 헤더
    # - 개별 질의응답(HEADER, sheet1)과는 별도 탭에 저장하여, "개별 답변 로그"와
    #   "여러 질문을 묶어 재구성한 종합 문서"가 한 시트에 섞이지 않도록 구분함.
    SUMMARY_SHEET_NAME = "종합문서"
    SUMMARY_HEADER = ["일시", "포함된질의건수", "종합문서요약", "종합문서전체", "사용자구분"]

    # 지식베이스 탭(워크시트) 이름 및 헤더
    # 설계 의도 (2026-06-25 추가):
    # - 로컬 _knowledge 폴더의 .txt 파일은 PC에서 직접 수정 후 GitHub에 다시 올려야만
    #   갱신되므로, 웹 서버에서 "지식베이스에 확정 저장"을 눌러도 그 효과가 서버
    #   재시작 시 사라지는 문제가 있었음(서버는 GitHub의 임시 복제본일 뿐이라서).
    # - 해결: 새로 확정되는 지식은 이 구글시트 "지식베이스" 탭에 쌓음. PC/웹 어디서
    #   저장하든 같은 구글시트에 들어가므로 항상 최신 상태로 동기화됨.
    # - 기존 .txt 파일들은 "최초 기본 지식"으로 그대로 두고, 양이 많아지면 이 탭의
    #   내용을 다운로드해서 .txt로 정리한 뒤 GitHub에 다시 올리는 방식으로 주기적으로
    #   "승격"시키는 운영 방식을 전제로 함.
    KNOWLEDGE_SHEET_NAME = "지식베이스"
    KNOWLEDGE_HEADER = ["일시", "분류", "질문", "확정내용"]

    # 검증대기 탭(워크시트) 이름 및 헤더
    # 설계 의도 (2026-06-27 추가 — WebSocket 연결 끊김으로 인한 작업 손실 방지):
    # - Streamlit Cloud는 브라우저-서버 간 WebSocket 연결이 간헐적으로 끊기면
    #   세션(session_state)이 통째로 초기화되는 알려진 플랫폼 특성이 있음. 이게
    #   "지식베이스에 확정 저장" 흐름 중간(검증 완료 후 PIN 입력 전)에 발생하면,
    #   회계사가 다시 로그인 → 검색기록에서 같은 항목 찾기 → 웹검색 재검증을
    #   처음부터 다시 돌려야 하는 큰 불편이 있었음.
    # - 해결: 검증(_run_verification_search + _apply_corrections)이 끝나는 즉시,
    #   그 결과(원본 질문, 수정된 최종본, 추천 파일, 수정 요약)를 이 탭에 1행으로
    #   자동 저장해둠. 화면이 튕겨도 사이드바의 "검증대기 불러오기"로 즉시 복원
    #   가능 — 웹검색을 다시 돌릴 필요 없이 PIN 입력 단계로 바로 진입함.
    # - 확정 저장이 완료되면 해당 행은 자동 삭제됨(완료된 항목이 계속 쌓이지
    #   않도록). 즉 이 탭은 "현재 진행 중인 작업"만 담는 임시 작업공간임.
    PENDING_SHEET_NAME = "검증대기"
    PENDING_HEADER = ["일시", "질문", "원본답변", "수정된내용", "수정요약", "추천파일", "추천이유"]

    # 교차검증대기 탭(워크시트) 이름 및 헤더
    # 설계 의도 (2026-06-28 추가 — 교차검증 스레드 단계의 WebSocket 끊김 대비):
    # - 위의 "검증대기" 탭은 ①②③ 자동검증+수정이 끝난 "확정 직전" 단계만
    #   백업함. 그런데 v1.6에서 새로 추가된 교차검증 스레드(1차 검증 →
    #   다른 AI에게 질문 → 답변 받아 재검증 → 2차, 3차... 반복)는 회계사가
    #   다른 AI 사이트로 탭을 옮겨 한참 머무는 구간이 많아, 정작 WebSocket이
    #   가장 자주 끊기는 지점인데도 백업이 전혀 없었음. 끊기면 진행 중이던
    #   모든 라운드 기록이 그대로 사라지는 문제가 있었음(회계사 피드백 반영).
    # - 해결: 매 라운드(재검증)가 끝날 때마다 그 라운드를 이 탭에 한 행씩
    #   추가함. 같은 작업(같은 검증 스레드)에서 나온 라운드들은 동일한
    #   세션ID로 묶여, 나중에 그 세션ID로 전체를 모아 복원할 수 있음.
    # - "검증대기" 탭과 별도 탭으로 분리한 이유: 컬럼 구조가 다르고(이쪽은
    #   라운드별로 여러 행), 기존에 이미 안정적으로 동작하는 "검증대기"
    #   탭의 컬럼 구조를 건드리지 않기 위함.
    # - 회계사가 "확정 단계로 진행"을 눌러 다음 단계로 넘어가거나, 확정
    #   저장/취소가 완료되면 해당 세션ID의 모든 행을 정리함 — 이 탭도
    #   "검증대기" 탭과 마찬가지로 "지금 진행 중인 작업"만 담는 임시 공간임.
    CROSSCHECK_SHEET_NAME = "교차검증대기"
    CROSSCHECK_HEADER = [
        "세션ID", "일시", "라운드", "질문", "원본답변",
        "검증결과", "보낸질문", "외부AI답변", "재검증후결과",
    ]

    # 계정설정 탭(워크시트) 이름 및 헤더
    # 설계 의도 (2026-06-27 추가 — 비밀번호를 .env/Secrets가 아닌 화면에서 관리):
    # - 관리자(회계사)/직원 로그인 비밀번호를 처음에는 .env(로컬) 또는 Streamlit
    #   Secrets(웹)의 ADMIN_PASSWORD/STAFF_PASSWORD로 관리했으나, 바꿀 때마다
    #   파일을 직접 열어 고치고 재배포해야 해서 번거롭다는 피드백을 받음.
    # - 해결: 비밀번호의 해시값을 이 구글시트 탭에 저장하고, 관리자가 로그인 후
    #   사이드바에서 직접 변경할 수 있게 함(PIN 변경과 동일한 패턴). 평문이
    #   아니라 sha256 해시만 저장함 — PIN 저장 방식과 동일한 보안 수준.
    # - 이 탭에 값이 하나도 없는 상태(앱을 처음 띄운 경우)에는, 로그인 화면 대신
    #   "최초 계정 설정" 화면을 보여줘 그 자리에서 관리자/직원 비밀번호를 처음
    #   만들게 함. 따라서 .env/Secrets에 비밀번호를 미리 적어둘 필요가 전혀
    #   없어짐. (이 화면은 인증 전 누구나 접근 가능하므로, 배포 직후 가능한
    #   빨리 설정을 완료하는 것을 전제로 함 — 약한 추가 보호장치는 의도적으로
    #   두지 않음. 그조차 별도 설정값이 필요해 번거로움을 다시 만들기 때문)
    ACCOUNT_SHEET_NAME = "계정설정"
    ACCOUNT_HEADER = ["역할", "비밀번호해시"]

    def __init__(self, sheet_id: str = None, credentials_path: str = None, credentials_json: str = None):
        """
        Parameters
        ----------
        sheet_id : str
            구글 스프레드시트 ID (없으면 .env/Secrets의 GOOGLE_SHEET_ID 사용)
        credentials_path : str
            서비스 계정 JSON 키 '파일 경로' (PC 로컬 환경에서 사용. 없으면
            .env의 GOOGLE_CREDENTIALS_PATH 사용)
        credentials_json : str
            서비스 계정 JSON 키 '내용 전체(텍스트)'. 웹 배포 환경(Streamlit Cloud의
            Secrets 등)에서는 PC의 파일 경로를 그대로 쓸 수 없으므로, JSON 파일 내용을
            문자열 그대로 이 값으로 전달함 (없으면 .env/Secrets의 GOOGLE_CREDENTIALS_JSON 사용).

        우선순위: credentials_json(텍스트)이 있으면 그걸 사용하고, 없으면
        credentials_path(파일 경로)를 사용함. 즉 PC에서는 기존처럼 파일 경로만
        설정해두면 그대로 동작하고, 웹 배포 환경에서는 GOOGLE_CREDENTIALS_JSON만
        설정해두면 동작함. 두 환경을 같은 코드로 모두 지원하기 위함.
        """
        self.enabled = False
        self.sheet = None
        self.summary_sheet = None
        self.knowledge_sheet = None
        self.pending_sheet = None
        self.crosscheck_sheet = None
        self.account_sheet = None
        self.error_message = ""

        sheet_id = (sheet_id or os.getenv("GOOGLE_SHEET_ID", "")).strip()
        credentials_path = (credentials_path or os.getenv("GOOGLE_CREDENTIALS_PATH", "")).strip()
        credentials_json = (credentials_json or os.getenv("GOOGLE_CREDENTIALS_JSON", "")).strip()

        if not sheet_id or not (credentials_path or credentials_json):
            # 설정이 없으면 조용히 비활성 상태로 둠 (선택 기능)
            self.error_message = "__NOT_CONFIGURED__"  # 진단용: "설정 자체가 없음"과 "설정했는데 실패"를 구분
            return

        try:
            import gspread

            if credentials_json:
                # 웹 배포 환경: JSON 텍스트를 그대로 사용 (파일을 거치지 않음)
                import json
                try:
                    credentials_dict = json.loads(credentials_json)
                except json.JSONDecodeError as e:
                    self.error_message = f"GOOGLE_CREDENTIALS_JSON 형식이 올바른 JSON이 아닙니다: {e}"
                    print(f"[경고] 구글 시트 로깅 비활성화: {self.error_message}")
                    return
                gc = gspread.service_account_from_dict(credentials_dict)
            else:
                # PC 로컬 환경: 파일 경로 방식 (기존 방식 그대로 유지)
                if not Path(credentials_path).exists():
                    self.error_message = f"인증키 파일을 찾을 수 없습니다: {credentials_path}"
                    print(f"[경고] 구글 시트 로깅 비활성화: {self.error_message}")
                    return
                gc = gspread.service_account(filename=credentials_path)

            spreadsheet = gc.open_by_key(sheet_id)
            self.sheet = spreadsheet.sheet1

            # 헤더가 없으면 추가
            existing = self.sheet.row_values(1)
            if existing != self.HEADER:
                self.sheet.insert_row(self.HEADER, 1)

            # 종합문서 탭 확인/생성 (없으면 새로 만듦)
            try:
                self.summary_sheet = spreadsheet.worksheet(self.SUMMARY_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                self.summary_sheet = spreadsheet.add_worksheet(
                    title=self.SUMMARY_SHEET_NAME, rows=200, cols=len(self.SUMMARY_HEADER)
                )

            existing_summary_header = self.summary_sheet.row_values(1)
            if existing_summary_header != self.SUMMARY_HEADER:
                self.summary_sheet.insert_row(self.SUMMARY_HEADER, 1)

            # 지식베이스 탭 확인/생성 (없으면 새로 만듦)
            try:
                self.knowledge_sheet = spreadsheet.worksheet(self.KNOWLEDGE_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                self.knowledge_sheet = spreadsheet.add_worksheet(
                    title=self.KNOWLEDGE_SHEET_NAME, rows=500, cols=len(self.KNOWLEDGE_HEADER)
                )

            existing_knowledge_header = self.knowledge_sheet.row_values(1)
            if existing_knowledge_header != self.KNOWLEDGE_HEADER:
                self.knowledge_sheet.insert_row(self.KNOWLEDGE_HEADER, 1)

            # 검증대기 탭 확인/생성 (없으면 새로 만듦)
            try:
                self.pending_sheet = spreadsheet.worksheet(self.PENDING_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                self.pending_sheet = spreadsheet.add_worksheet(
                    title=self.PENDING_SHEET_NAME, rows=50, cols=len(self.PENDING_HEADER)
                )

            existing_pending_header = self.pending_sheet.row_values(1)
            if existing_pending_header != self.PENDING_HEADER:
                self.pending_sheet.insert_row(self.PENDING_HEADER, 1)

            # 교차검증대기 탭 확인/생성 (없으면 새로 만듦)
            try:
                self.crosscheck_sheet = spreadsheet.worksheet(self.CROSSCHECK_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                self.crosscheck_sheet = spreadsheet.add_worksheet(
                    title=self.CROSSCHECK_SHEET_NAME, rows=50, cols=len(self.CROSSCHECK_HEADER)
                )

            existing_crosscheck_header = self.crosscheck_sheet.row_values(1)
            if existing_crosscheck_header != self.CROSSCHECK_HEADER:
                self.crosscheck_sheet.insert_row(self.CROSSCHECK_HEADER, 1)

            # 계정설정 탭 확인/생성 (없으면 새로 만듦)
            try:
                self.account_sheet = spreadsheet.worksheet(self.ACCOUNT_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                self.account_sheet = spreadsheet.add_worksheet(
                    title=self.ACCOUNT_SHEET_NAME, rows=10, cols=len(self.ACCOUNT_HEADER)
                )

            existing_account_header = self.account_sheet.row_values(1)
            if existing_account_header != self.ACCOUNT_HEADER:
                self.account_sheet.insert_row(self.ACCOUNT_HEADER, 1)

            self.enabled = True
            self.error_message = ""  # 성공했으므로 명시적으로 비움
        except ImportError as e:
            self.error_message = f"[ImportError] gspread/google-auth 패키지 문제: {e}"
            print(f"[경고] 구글 시트 로깅 비활성화: {self.error_message}")
        except Exception as e:
            # 진단용: 예외 타입명을 반드시 포함시켜서, error_message가 비어 보이는 일이 없게 함
            self.error_message = f"[{type(e).__name__}] {e}"
            print(f"[경고] 구글 시트 로깅 초기화 실패: {self.error_message}")

    def log(self, question: str, answer: str, evidence_type: str = "", user_type: str = "직원") -> bool:
        """
        질의응답 1건을 시트에 한 행으로 추가.

        Parameters
        ----------
        question : str
            사용자 질문
        answer : str
            AI 답변 전체
        evidence_type : str
            "1순위(내부)", "2순위(법령)", "3순위(예규검색)", "4순위(AI일반지식)" 등 요약
        user_type : str
            "회계사" 또는 "직원" (구분 표시용, 로그인 시스템이 없으므로 단순 라벨)

        Returns
        -------
        bool
            성공 여부 (실패해도 예외를 던지지 않음)
        """
        if not self.enabled:
            return False

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            summary = answer[:100].replace("\n", " ") + ("..." if len(answer) > 100 else "")
            self.sheet.append_row(
                [timestamp, question, summary, answer, evidence_type, user_type],
                value_input_option="RAW",
            )
            return True
        except Exception as e:
            print(f"[경고] 구글 시트 로깅 실패 (답변 생성에는 영향 없음): {e}")
            return False

    def get_recent_logs(self, limit: int = 20) -> list:
        """
        최근 로그를 가져옴 (시트 내 검색/조회 UI에서 사용)

        Returns
        -------
        list[dict]
            실패 시 빈 리스트
        """
        if not self.enabled:
            return []

        try:
            all_rows = self.sheet.get_all_records()
            return all_rows[-limit:][::-1]  # 최신순
        except Exception as e:
            print(f"[경고] 구글 시트 조회 실패: {e}")
            return []

    def delete_log(self, timestamp: str, question: str) -> bool:
        """
        개별 질의응답 기록 1건을 시트에서 삭제.

        설계 의도 (2026-06-25 추가):
        - 검색 기록 화면에서 회계사가 PIN을 입력해야만 삭제 가능하도록 UI 쪽에서
          PIN 확인을 먼저 거침 (이 메서드 자체는 PIN을 검증하지 않음 — 호출하는
          쪽(streamlit_ui.py)에서 PIN 확인 후에만 이 메서드를 부르는 책임을 짐).
        - 행을 정확히 식별하기 위해 '일시'와 '질문' 두 값이 모두 일치하는 행을 찾음.
          일시(타임스탬프)는 초 단위까지 기록되므로 거의 항상 고유하지만, 혹시
          같은 초에 같은 질문이 중복 기록된 극단적인 경우까지 고려해 두 값을
          모두 대조함. 일치하는 첫 번째 행만 삭제함(완전히 동일한 행이 여러 개
          있어도 하나만 지워짐 — 이런 경우는 실질적으로 거의 없음).

        Parameters
        ----------
        timestamp : str
            삭제할 행의 '일시' 값 (시트에 기록된 그대로, 예: "2026-06-25 02:36:34")
        question : str
            삭제할 행의 '질문' 값 (timestamp와 함께 행을 정확히 식별하기 위함)

        Returns
        -------
        bool
            삭제 성공 여부 (해당 행을 못 찾은 경우도 False)
        """
        if not self.enabled:
            return False

        try:
            all_values = self.sheet.get_all_values()  # 헤더 포함, 1행부터
            for row_idx, row in enumerate(all_values[1:], start=2):  # 시트 행 번호는 1부터, 헤더 제외하고 2행부터
                if len(row) >= 2 and row[0] == timestamp and row[1] == question:
                    self.sheet.delete_rows(row_idx)
                    return True
            return False  # 일치하는 행을 못 찾음
        except Exception as e:
            print(f"[경고] 구글 시트 행 삭제 실패: {e}")
            return False

    # ------------------------------------------------------------------
    # 종합문서 전용 기록/조회 (개별 질의응답과는 별도 탭에 저장)
    # ------------------------------------------------------------------
    def log_summary(self, turns_count: int, summary_text: str, user_type: str = "직원") -> bool:
        """
        종합 문서 1건을 '종합문서' 탭에 한 행으로 추가.

        주의: 개별 질의응답(log 메서드)과는 달리, 이 메서드는 사용자가 화면에서
        "종합 문서 저장" 버튼을 직접 눌렀을 때만 호출됨 (자동 기록 아님).
        매번 종합 문서를 만들어볼 때마다 기록하면 시험적으로 눌러본 것까지 다 쌓여
        실제로 저장해 둔 것과 구분이 안 되기 때문.

        Parameters
        ----------
        turns_count : int
            이 종합 문서에 포함된 개별 질의응답 건수
        summary_text : str
            AI가 재구성한 종합 문서 전체 내용
        user_type : str
            "회계사" 또는 "직원"

        Returns
        -------
        bool
            성공 여부 (실패해도 예외를 던지지 않음)
        """
        if not self.enabled or self.summary_sheet is None:
            return False

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            summary_preview = summary_text[:100].replace("\n", " ") + ("..." if len(summary_text) > 100 else "")
            self.summary_sheet.append_row(
                [timestamp, turns_count, summary_preview, summary_text, user_type],
                value_input_option="RAW",
            )
            return True
        except Exception as e:
            print(f"[경고] 종합문서 시트 로깅 실패 (저장 자체에는 영향 없음): {e}")
            return False

    def get_recent_summaries(self, limit: int = 20) -> list:
        """
        '종합문서' 탭에서 최근 기록을 가져옴 (조회 UI에서 사용)

        Returns
        -------
        list[dict]
            실패 시 빈 리스트
        """
        if not self.enabled or self.summary_sheet is None:
            return []

        try:
            all_rows = self.summary_sheet.get_all_records()
            return all_rows[-limit:][::-1]  # 최신순
        except Exception as e:
            print(f"[경고] 종합문서 시트 조회 실패: {e}")
            return []

    def delete_summary(self, timestamp: str) -> bool:
        """
        종합 문서 기록 1건을 '종합문서' 탭에서 삭제.

        개별 질의응답(delete_log)과 동일한 설계 원칙을 따름: PIN 확인은 호출하는
        쪽(streamlit_ui.py)의 책임이며, 이 메서드는 삭제 동작만 수행함.
        종합문서는 '일시'만으로도 충분히 고유하므로(같은 초에 두 번 생성하기 어려움)
        일시 하나만으로 행을 식별함.

        Parameters
        ----------
        timestamp : str
            삭제할 행의 '일시' 값

        Returns
        -------
        bool
            삭제 성공 여부 (해당 행을 못 찾은 경우도 False)
        """
        if not self.enabled or self.summary_sheet is None:
            return False

        try:
            all_values = self.summary_sheet.get_all_values()
            for row_idx, row in enumerate(all_values[1:], start=2):
                if len(row) >= 1 and row[0] == timestamp:
                    self.summary_sheet.delete_rows(row_idx)
                    return True
            return False
        except Exception as e:
            print(f"[경고] 종합문서 시트 행 삭제 실패: {e}")
            return False


    # ------------------------------------------------------------------
    # 지식베이스 (구글시트 기반) - 회계사 확정 저장한 새 지식을 누적 보관
    # ------------------------------------------------------------------
    def add_knowledge_entry(self, category: str, question: str, confirmed_content: str) -> bool:
        """
        회계사가 확정한 새 지식 1건을 '지식베이스' 탭에 한 행으로 추가.

        주의: 이 메서드는 PIN 검증을 하지 않음 — 호출하는 쪽(streamlit_ui.py)에서
        verify_pin()으로 확인한 뒤에만 호출해야 함.

        Parameters
        ----------
        category : str
            분류 (기존 .txt 파일명과 호환되는 분류명을 권장하되, 자유 텍스트도 가능.
            예: "01_공통_세무질의회신집", "부가세", "법인세" 등)
        question : str
            원래 질문
        confirmed_content : str
            회계사가 검토/수정한 최종 확정 내용

        Returns
        -------
        bool
            성공 여부
        """
        if not self.enabled or self.knowledge_sheet is None:
            return False

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.knowledge_sheet.append_row(
                [timestamp, category, question, confirmed_content],
                value_input_option="RAW",
            )
            return True
        except Exception as e:
            print(f"[경고] 지식베이스 시트 추가 실패: {e}")
            return False

    def get_all_knowledge_text(self) -> str:
        """
        '지식베이스' 탭에 쌓인 모든 항목을 하나의 텍스트로 합쳐서 반환.
        AI 답변 생성 시 로컬 .txt 파일 내용과 합쳐서 참고 자료로 사용함.

        Returns
        -------
        str
            모든 항목을 정리한 텍스트. 항목이 없거나 실패 시 빈 문자열 반환
            (빈 문자열이면 호출하는 쪽에서 이 부분을 그냥 건너뛰면 됨 — 안전한 기본값).
        """
        if not self.enabled or self.knowledge_sheet is None:
            return ""

        try:
            rows = self.knowledge_sheet.get_all_records()
        except Exception as e:
            print(f"[경고] 지식베이스 시트 조회 실패: {e}")
            return ""

        if not rows:
            return ""

        chunks = []
        for row in rows:
            chunks.append(
                f"[분류: {row.get('분류', '')}] (확정일시: {row.get('일시', '')})\n"
                f"질의: {row.get('질문', '')}\n"
                f"확정내용: {row.get('확정내용', '')}\n"
                f"---"
            )
        return "\n\n".join(chunks)

    def get_recent_knowledge_entries(self, limit: int = 50) -> list:
        """
        '지식베이스' 탭의 최근 항목을 가져옴 (조회/관리 UI에서 사용).

        Returns
        -------
        list[dict]
            실패 시 빈 리스트
        """
        if not self.enabled or self.knowledge_sheet is None:
            return []

        try:
            all_rows = self.knowledge_sheet.get_all_records()
            return all_rows[-limit:][::-1]  # 최신순
        except Exception as e:
            print(f"[경고] 지식베이스 시트 조회 실패: {e}")
            return []


    # ------------------------------------------------------------------
    # 검증대기 (구글시트 기반) - WebSocket 끊김으로 작업 중단 시 복구용 임시저장
    # ------------------------------------------------------------------
    def save_pending_verification(
        self,
        question: str,
        original_answer: str,
        corrected_content: str,
        correction_summary: str,
        recommended_file: str,
        recommended_reason: str,
    ) -> str:
        """
        검증/자동수정이 끝난 결과를 '검증대기' 탭에 한 행으로 저장.

        웹검색 재검증("① 내용 검증하기")이 끝나는 즉시 호출해서, 그 결과를
        구글시트에 즉시 백업해둠. 이후 PIN 입력 전에 화면이 튕겨도(WebSocket
        연결 끊김 등), 사이드바에서 이 행을 불러와 검증을 다시 돌리지 않고
        바로 PIN 입력 단계로 이어갈 수 있음.

        Returns
        -------
        str
            저장된 행의 식별자로 쓸 타임스탬프 문자열. 저장 실패 시 빈 문자열.
        """
        if not self.enabled or self.pending_sheet is None:
            return ""

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.pending_sheet.append_row(
                [
                    timestamp,
                    question,
                    original_answer,
                    corrected_content,
                    correction_summary,
                    recommended_file,
                    recommended_reason,
                ],
                value_input_option="RAW",
            )
            return timestamp
        except Exception as e:
            print(f"[경고] 검증대기 시트 저장 실패: {e}")
            return ""

    def list_pending_verifications(self) -> list:
        """
        '검증대기' 탭에 남아있는 모든 항목을 가져옴 (최신순).
        사이드바에 "이전에 검증해두고 못 저장한 항목" 목록을 보여줄 때 사용.

        Returns
        -------
        list[dict]
            실패하거나 항목이 없으면 빈 리스트
        """
        if not self.enabled or self.pending_sheet is None:
            return []

        try:
            all_rows = self.pending_sheet.get_all_records()
            return all_rows[::-1]  # 최신순
        except Exception as e:
            print(f"[경고] 검증대기 시트 조회 실패: {e}")
            return []

    def delete_pending_verification(self, timestamp: str) -> bool:
        """
        '검증대기' 탭에서 특정 일시(timestamp)에 해당하는 행을 삭제.
        확정 저장이 완료된 직후, 더 이상 필요 없는 임시저장 행을 정리하기 위해 호출.

        Returns
        -------
        bool
            성공 여부 (해당 행을 못 찾아도 예외 없이 False만 반환)
        """
        if not self.enabled or self.pending_sheet is None:
            return False

        try:
            all_values = self.pending_sheet.get_all_values()
            for row_idx, row in enumerate(all_values[1:], start=2):
                if len(row) >= 1 and row[0] == timestamp:
                    self.pending_sheet.delete_rows(row_idx)
                    return True
            return False
        except Exception as e:
            print(f"[경고] 검증대기 시트 행 삭제 실패: {e}")
            return False

    # ------------------------------------------------------------------
    # 교차검증대기 (구글시트 기반) - 교차검증 스레드(1차/2차/3차...) 단계의
    # WebSocket 끊김 대비 임시저장. "검증대기"와 똑같은 목적이지만, 컬럼
    # 구조가 다르고(여러 라운드가 여러 행으로 쌓임) 대상 단계가 다름(②번
    # 교차검증 단계 — 다른 AI 사이트로 탭을 옮겨 머무는 일이 많아 가장 자주
    # 끊기는 구간).
    # ------------------------------------------------------------------
    def save_crosscheck_round(
        self,
        session_id: str,
        round_no: int,
        question: str,
        original_answer: str,
        verification_text: str,
        cross_prompt: str,
        external_ai_input: str,
        next_verification_text: str = "",
    ) -> bool:
        """
        교차검증 스레드의 한 라운드(예: 1차)가 완료될 때마다 호출해서, 그
        라운드 내용을 '교차검증대기' 탭에 새 행으로 추가함.

        같은 작업(같은 검증 스레드)에서 나온 라운드들은 모두 동일한
        session_id를 가지므로, 나중에 get_crosscheck_rounds(session_id)로
        한꺼번에 모아서 복원할 수 있음.

        Parameters
        ----------
        session_id : str
            이 검증 스레드를 식별하는 고유 ID. render_confirm_to_kb_button에서
            "지식베이스에 확정 저장" 버튼을 누를 때 한 번 생성해서, 같은
            작업의 모든 라운드에 동일하게 사용함.
        round_no : int
            이번에 완료된 라운드 번호 (1, 2, 3...)
        question, original_answer : str
            원래 질문과 검증 대상 원본 답변 (라운드마다 동일한 값 — 복원 시
            첫 행에서만 읽어도 되지만, 행마다 같이 저장해 단순하게 둠)
        verification_text : str
            이 라운드가 "시작될 때" 보여줬던 검증 결과 텍스트 (즉 외부 AI에게
            보낼 질문을 만드는 데 쓰인 텍스트)
        cross_prompt : str
            이 라운드에서 다른 AI에게 보냈던 질문 문구
        external_ai_input : str
            이 라운드에서 외부 AI로부터 받아온 답변
        next_verification_text : str, optional
            이 라운드의 외부 AI 답변까지 반영해 재검증한 "다음 결과" 텍스트.
            버그 수정 (2026-06-28 — 교차검증대기 복원 시 다음 라운드를 다시
            보여주지 못하던 문제): 이 값이 없으면, 복원했을 때 "지금 진행
            중인 라운드"가 무엇을 보여줘야 할지 알 수 없어, 이미 끝난
            라운드의 검증 결과를 다시 보여주는 부정확한 복원이 됨. 이 값을
            함께 저장해두면, 복원 시 정확히 "재검증까지 끝난 최신 결과"부터
            이어서 보여줄 수 있음.

        Returns
        -------
        bool
            성공 여부
        """
        if not self.enabled or self.crosscheck_sheet is None:
            return False

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.crosscheck_sheet.append_row(
                [
                    session_id,
                    timestamp,
                    round_no,
                    question,
                    original_answer,
                    verification_text,
                    cross_prompt,
                    external_ai_input,
                    next_verification_text,
                ],
                value_input_option="RAW",
            )
            return True
        except Exception as e:
            print(f"[경고] 교차검증대기 시트 저장 실패: {e}")
            return False

    def list_crosscheck_sessions(self) -> list:
        """
        '교차검증대기' 탭에 남아있는 세션들의 요약 목록을 가져옴 (최신순,
        세션당 1건 — 가장 최근 라운드 기준).
        사이드바에 "진행 중인 교차검증 불러오기" 목록을 보여줄 때 사용.

        Returns
        -------
        list[dict]
            각 항목: {"session_id", "question", "latest_round", "updated_at"}
            실패하거나 항목이 없으면 빈 리스트
        """
        if not self.enabled or self.crosscheck_sheet is None:
            return []

        try:
            all_rows = self.crosscheck_sheet.get_all_records()
            sessions = {}
            for row in all_rows:
                sid = row.get("세션ID", "")
                if not sid:
                    continue
                # 같은 세션ID의 행 중 라운드 번호가 가장 큰(=가장 최근) 것만 대표로 남김
                prev = sessions.get(sid)
                if prev is None or int(row.get("라운드", 0) or 0) >= int(prev["latest_round"]):
                    sessions[sid] = {
                        "session_id": sid,
                        "question": row.get("질문", ""),
                        "latest_round": int(row.get("라운드", 0) or 0),
                        "updated_at": row.get("일시", ""),
                    }
            return sorted(sessions.values(), key=lambda x: x["updated_at"], reverse=True)
        except Exception as e:
            print(f"[경고] 교차검증대기 세션 목록 조회 실패: {e}")
            return []

    def get_crosscheck_rounds(self, session_id: str) -> list:
        """
        특정 세션ID에 해당하는 모든 라운드 행을 라운드 순서대로 가져옴.
        화면의 verification_thread 형태로 그대로 복원하기 위한 용도.

        Returns
        -------
        list[dict]
            각 항목: {"round", "question", "original_answer",
                      "verification_text", "cross_prompt", "external_ai_input"}
            실패하거나 항목이 없으면 빈 리스트
        """
        if not self.enabled or self.crosscheck_sheet is None:
            return []

        try:
            all_rows = self.crosscheck_sheet.get_all_records()
            rounds = [
                {
                    "round": int(row.get("라운드", 0) or 0),
                    "question": row.get("질문", ""),
                    "original_answer": row.get("원본답변", ""),
                    "verification_text": row.get("검증결과", ""),
                    "cross_prompt": row.get("보낸질문", ""),
                    "external_ai_input": row.get("외부AI답변", ""),
                    "next_verification_text": row.get("재검증후결과", ""),
                }
                for row in all_rows
                if row.get("세션ID", "") == session_id
            ]
            rounds.sort(key=lambda r: r["round"])
            return rounds
        except Exception as e:
            print(f"[경고] 교차검증대기 세션 조회 실패: {e}")
            return []

    def delete_crosscheck_session(self, session_id: str) -> bool:
        """
        특정 세션ID에 해당하는 모든 행을 '교차검증대기' 탭에서 삭제함.
        회계사가 "확정 단계로 진행"을 누르거나, 확정 저장/취소가 완료되면
        호출해서, 더 이상 필요 없는 임시저장 행들을 한꺼번에 정리함.

        Returns
        -------
        bool
            하나 이상 삭제했으면 True, 대상이 없거나 실패하면 False
        """
        if not self.enabled or self.crosscheck_sheet is None:
            return False

        try:
            all_values = self.crosscheck_sheet.get_all_values()
            # 뒤에서부터 삭제해야 앞쪽 행 삭제로 인한 인덱스 밀림이 안 생김
            rows_to_delete = [
                row_idx
                for row_idx, row in enumerate(all_values[1:], start=2)
                if len(row) >= 1 and row[0] == session_id
            ]
            for row_idx in reversed(rows_to_delete):
                self.crosscheck_sheet.delete_rows(row_idx)
            return len(rows_to_delete) > 0
        except Exception as e:
            print(f"[경고] 교차검증대기 세션 삭제 실패: {e}")
            return False

    # ------------------------------------------------------------------
    # 계정설정 (구글시트 기반) - 관리자/직원 로그인 비밀번호를 화면에서 관리
    # ------------------------------------------------------------------
    def is_account_setup_done(self) -> bool:
        """
        '계정설정' 탭에 관리자/직원 비밀번호가 둘 다 설정되어 있는지 확인.
        하나라도 없으면 "최초 계정 설정" 화면을 보여줘야 함을 의미함.

        Returns
        -------
        bool
            True면 이미 설정 완료. 구글시트 로깅이 비활성 상태면 항상 False를
            반환함(이 경우 호출하는 쪽에서 .env/Secrets 기반 방식으로 폴백해야 함).
        """
        if not self.enabled or self.account_sheet is None:
            return False

        try:
            rows = self.account_sheet.get_all_records()
            roles_set = {r.get("역할") for r in rows if r.get("비밀번호해시")}
            return "admin" in roles_set and "staff" in roles_set
        except Exception as e:
            print(f"[경고] 계정설정 시트 조회 실패: {e}")
            return False

    def _get_password_hash(self, role: str) -> str:
        """role("admin" 또는 "staff")에 해당하는 저장된 비밀번호 해시를 반환. 없으면 빈 문자열."""
        if not self.enabled or self.account_sheet is None:
            return ""

        try:
            rows = self.account_sheet.get_all_records()
            for r in rows:
                if r.get("역할") == role:
                    return str(r.get("비밀번호해시", ""))
            return ""
        except Exception as e:
            print(f"[경고] 계정설정 시트 조회 실패: {e}")
            return ""

    def verify_account_password(self, role: str, input_password: str) -> bool:
        """입력한 비밀번호가 해당 role에 저장된 해시와 일치하는지 확인."""
        stored_hash = self._get_password_hash(role)
        if not stored_hash:
            return False
        input_hash = hashlib.sha256(input_password.strip().encode("utf-8")).hexdigest()
        return stored_hash == input_hash

    def set_account_password(self, role: str, new_password: str) -> bool:
        """
        role("admin" 또는 "staff")의 비밀번호를 설정/변경. 평문이 아니라
        sha256 해시만 저장함(PIN 저장 방식과 동일한 보안 수준).

        기존에 그 역할의 행이 있으면 해시값만 덮어쓰고, 없으면 새 행을 추가함.

        Returns
        -------
        bool
            성공 여부
        """
        if not self.enabled or self.account_sheet is None:
            return False

        try:
            new_hash = hashlib.sha256(new_password.strip().encode("utf-8")).hexdigest()
            all_values = self.account_sheet.get_all_values()
            for row_idx, row in enumerate(all_values[1:], start=2):
                if len(row) >= 1 and row[0] == role:
                    self.account_sheet.update_cell(row_idx, 2, new_hash)
                    return True
            # 기존 행이 없으면 새로 추가
            self.account_sheet.append_row([role, new_hash], value_input_option="RAW")
            return True
        except Exception as e:
            print(f"[경고] 계정설정 시트 저장 실패: {e}")
            return False


# ----------------------------------------------------------------------
# 단독 실행 테스트
# ----------------------------------------------------------------------
if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()

    print("=" * 60)
    print("sheet_logger.py 단독 실행 테스트")
    print("=" * 60)

    logger = SheetLogger()
    if not logger.enabled:
        print(f"\n로깅이 비활성화되어 있습니다.")
        if logger.error_message:
            print(f"원인: {logger.error_message}")
        else:
            print("원인: .env에 GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_PATH가 설정되지 않음")
    else:
        print("\n로깅 활성화됨. 테스트 행을 추가합니다...")
        success = logger.log(
            question="[테스트] sheet_logger.py 단독 실행 테스트입니다",
            answer="이것은 연동 테스트용 답변입니다.",
            evidence_type="테스트",
            user_type="시스템",
        )
        print(f"기록 성공: {success}")

        print("\n최근 로그 5건:")
        for row in logger.get_recent_logs(5):
            print(f"  - {row}")
