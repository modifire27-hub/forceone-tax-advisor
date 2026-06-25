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
