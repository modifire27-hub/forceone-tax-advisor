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
