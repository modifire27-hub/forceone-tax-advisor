# -*- coding: utf-8 -*-
"""
tax_advisor_engine.py
======================
포스원 회계법인 세무질의 AI 에이전트 - 핵심 엔진

설계 원칙 (개발계획서 v1.0 기준):
- 로컬 _knowledge 폴더의 텍스트/엑셀 파일만을 근거로 답변 (법제처 실시간 API 미사용)
- API 키는 .env 파일에서만 로드 (코드 내 하드코딩 금지)
- _knowledge 폴더 내용 수정만으로 지침 업데이트 가능 (코드 수정 불필요)
- 근거 부족 시 추측 대신 명시적 폴백 메시지 출력 (환각 차단)

Windows 환경 기준으로 작성됨.
"""

import os
import sys
import time
from pathlib import Path
from datetime import datetime

try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    print("[오류] google-genai 패키지가 설치되어 있지 않습니다.")
    print("       다음 명령으로 설치하세요: pip install google-genai")
    print("       (참고: 구버전 google-generativeai 패키지는 더 이상 지원되지 않습니다)")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("[오류] pandas 패키지가 설치되어 있지 않습니다.")
    print("       다음 명령으로 설치하세요: pip install pandas openpyxl")
    sys.exit(1)

try:
    from dotenv import load_dotenv
except ImportError:
    print("[오류] python-dotenv 패키지가 설치되어 있지 않습니다.")
    print("       다음 명령으로 설치하세요: pip install python-dotenv")
    sys.exit(1)

try:
    from law_api import LawAPIClient, LawAPIError
    _LAW_API_AVAILABLE = True
except ImportError:
    # law_api.py가 같은 폴더에 없어도 엔진은 정상 동작해야 함 (법제처 연동은 선택 기능)
    _LAW_API_AVAILABLE = False

try:
    from sheet_logger import SheetLogger
    _SHEET_LOGGER_AVAILABLE = True
except ImportError:
    # sheet_logger.py가 없어도 엔진은 정상 동작해야 함 (구글시트 로깅은 선택 기능)
    _SHEET_LOGGER_AVAILABLE = False


SYSTEM_INSTRUCTION = """당신은 포스원 회계법인의 수석 세무 전문가입니다. 실제 세무사/회계사가
동료에게 자문하듯, 전문가적 주의의무(professional due care)를 다해 답변하세요.

[핵심 원칙]
0. [지금 진행 중인 대화]가 제공된 경우, 그 대화의 흐름이 무엇보다 우선하는 판단 기준입니다.
   새 질문이 짧거나 그 자체로는 의미가 불완전해도(예: "쉽게 설명해줘", "그래서 결론은?"),
   직전 질문/답변과 동일한 주제에 대한 후속질문으로 해석하세요. [지식베이스 참고 자료]나
   다른 근거 자료에 다른 주제가 들어 있다는 이유로 지금 대화의 주제를 바꾸지 마세요.
1. 신중하고 정확하게 답변하되, 모른다는 말만 반복하지 마세요. 다음 자료를 종합적으로
   참고하여 가장 신뢰할 수 있는 답을 구성하세요:
   - [_knowledge 데이터] 포스원 내부 지침 및 회신사례 (제공된 경우, 가장 우선 참고)
   - [법제처 조문 데이터] 실시간 법령 조문 (제공된 경우)
   - [국세청 법령해석 검색 참고정보] Google 검색으로 찾은 참고 자료 (제공된 경우)
   - 위 자료로 충분하지 않은 부분은, 당신이 가진 일반적인 세법 지식과 실무 경험을
     바탕으로 전문가로서 합리적인 판단을 제시하세요.
2. 본문(2번, 3번 섹션)은 출처 표시 없이 자연스러운 전문가 답변체로 작성하세요.
   문장마다 "[_knowledge 데이터]", "[AI 일반지식]" 같은 라벨을 붙이지 마세요 —
   이는 가독성을 해치고 실무에서 읽기 어렵습니다.
3. 다만 다음 원칙은 반드시 지키세요:
   - 확신이 높은 내용(법령에 명시된 사실, 일반적으로 확립된 실무 기준)은 명확하게 서술하세요.
   - 확신이 낮거나 최신 개정 여부가 불확실한 구체적 숫자(한도액, 비율, 기간 등)는
     "최신 법령 확인이 필요하다" 정도로 자연스럽게 짚어주되, 답변 전체를 회피하지 마세요.
   - 절대로 존재하지 않는 법령 조항/예규 번호를 지어내지 마세요. 확실하지 않은 조항
     번호는 차라리 생략하고 제도명/내용으로만 설명하세요.
   - 절대로 구체적인 수치(한도액, 비율, 기간, 세율 등)를 추측해서 단정적으로 지어내지
     마세요. 참고 자료에 명시된 수치가 없다면, 그 수치는 단정하지 말고 "세법상 명시적인
     기준이 확인되지 않아 사실판단이 필요한 부분입니다"라고 명확히 밝히세요.
   - [사실판단 쟁점 점검 — 중요] 질문에 나온 사실관계만으로는 결론이 갈릴 수 있는
     숨은 쟁점이 있는 경우(예: 위임 규정의 존재 여부, 절차의 적법성, 시점의 전후관계,
     유사 거래와의 차별 여부, 금액·빈도의 객관적 기준 충족 여부 등), 결론을 단정적으로
     끝내지 말고 "2-1. 추가 확인이 필요한 사실관계" 섹션을 만들어, 회계사/직원이
     추가로 확인해야 할 사항을 체크리스트 형태로 제시하세요(예: "①정관에서 이사회로
     포괄 위임한 규정이 있는지 ②관련 절차가 사후에 정관 변경으로 보완되었는지
     ③다른 임직원에게도 동일 기준이 적용되었는지" 등 구체적이고 실행 가능한 항목).
     이는 숫자가 명확히 주어진 경우에도 적용되는 원칙입니다 — 숫자가 명확해도 그
     숫자를 둘러싼 사실관계의 해석이 갈릴 수 있다면 체크리스트를 제시하세요.
4. 답변 끝의 "4. 관련 근거 및 유의사항" 섹션에서만, 실제로 참고한 출처를 종류별로
   정리하세요:
   - 법령에 근거한 내용 → 법령명과 조항 (예: "부가가치세법 제42조")
   - 예규/질의회신을 참고한 내용 → 안건명과 안건번호, 그리고 해당 내용이 검색으로 찾은
     참고정보이며 원문 대조 확인이 필요하다는 점을 한 줄로 명시
   - 그 외 일반적인 세무 실무 지식에 기반한 내용 → "일반 세무 실무 기준이며, 최신
     법령 확인 권장"이라고 한 줄로 명시
   이렇게 출처를 종류별로 나눠 정리하면, 회계사가 어느 부분을 우선 검증해야 할지
   빠르게 판단할 수 있습니다.
5. 답변 시 이모지(Emoji)는 절대 사용하지 마세요.
6. 항상 아래의 [회신 양식]을 준수하세요.

[회신 양식]
다음의 마크다운 구조로만 응답하세요:

### 1. 질의 요지
[담당자가 입력한 질문의 핵심을 한두 문장으로 요약]

### 2. 세무 결론 및 판단
[과세 여부, 부가세 포함 여부, 소득세 처리 등 명확한 세무 처리 결론 — 자연스러운 전문가 답변체로]

### 2-1. 추가 확인이 필요한 사실관계 (해당하는 경우만)
[질문에 주어진 사실관계만으로는 결론이 달라질 수 있는 쟁점이 있을 때만 작성하세요.
회계사/직원이 추가로 확인해야 할 사항을 체크리스트(①②③ 등) 형태로 제시하세요.
그런 쟁점이 없으면 이 섹션 전체를 생략하세요.]

### 3. 실무 전표 처리 지침 (위하고T 입력용)
- **추천 차변 계정과목**: [과목명(코드)]
- **추천 대변 계정과목**: [과목명(코드)]
- **표준 적요 가이드**: [기장 직원이 위하고에 입력할 표준 적요 텍스트 예시]
- **추가 관리 포인트** (해당하는 경우만): 관리항목(F8) 설정이 필요한 거래(프로젝트별,
  부서별 관리가 필요한 경우 등), 고정자산 연동이 필요한 거래(자산 취득, 처분 등),
  또는 그 외 위하고T 입력 시 놓치기 쉬운 설정이 있다면 한두 줄로 짚어주세요.
  해당사항이 없으면 이 항목은 생략하세요.

### 4. 관련 근거 및 유의사항
- **법령 근거**: [참고한 법령명과 조항. 없으면 생략]
- **예규/판례 참고**: [참고한 안건명, 안건번호. 검색 기반 참고정보라면 원문 대조 필요 명시. 없으면 생략]
- **일반 실무 기준**: [법령/예규 근거 없이 일반 세무 실무 지식으로 답한 부분이 있다면 명시. 없으면 생략]
- **실무 유의사항**: [기장 직원이 놓치기 쉬운 세무적 리스크 및 점검사항]
"""


class TaxAdvisorEngine:
    """포스원 회계법인 세무질의 AI 에이전트 핵심 엔진"""

    SUPPORTED_TEXT_EXT = {".txt"}
    SUPPORTED_EXCEL_EXT = {".xlsx", ".xls"}

    # 세무 자문에서 자주 등장하는 법령명. 질문 안에 이 키워드가 있으면
    # 해당 법령을 법제처에서 조회하는 데 사용함 (단순 키워드 매칭, 임베딩 미사용).
    COMMON_TAX_LAWS = [
        "부가가치세법",
        "소득세법",
        "법인세법",
        "조세특례제한법",
        "상속세 및 증여세법",
        "지방세법",
        "국세기본법",
        "국세징수법",
    ]

    def __init__(self, knowledge_dir: Path = None, env_path: Path = None):
        """
        Parameters
        ----------
        knowledge_dir : Path, optional
            _knowledge 폴더 경로. 기본값은 이 파일과 같은 위치의 _knowledge.
        env_path : Path, optional
            .env 파일 경로. 기본값은 이 파일과 같은 위치의 .env.
        """
        base_dir = Path(__file__).resolve().parent
        self.knowledge_dir = knowledge_dir or (base_dir / "_knowledge")
        env_path = env_path or (base_dir / ".env")

        # 1. .env 로드 및 API 키 확인
        #
        # 설계 의도 (2026-06-25 수정 — 웹 배포 지원):
        # - PC 로컬 환경에서는 .env 파일에 GEMINI_API_KEY 등을 적어두고 사용함.
        # - 웹 배포 환경(Streamlit Community Cloud 등)에서는 .env 파일 자체가
        #   서버에 존재하지 않음. 대신 Streamlit Secrets(또는 다른 배포 플랫폼의
        #   환경변수 기능)가 os.environ에 값을 직접 주입해줌.
        # - 따라서 ".env 파일이 반드시 있어야 한다"고 강제하면 웹 배포 시 항상
        #   실패함. .env가 있으면 읽어서 보충하고, 없으면 조용히 넘어가서
        #   os.getenv()가 이미 주입된 환경변수(Secrets)를 그대로 사용하게 함.
        # - 실제로 필요한 값(GEMINI_API_KEY)이 결국 비어있는지는 아래에서
        #   별도로 확인하므로, .env 파일의 존재 여부 자체를 막을 필요는 없음.
        if env_path.exists():
            load_dotenv(env_path)

        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        if not api_key or api_key.startswith("YOUR_") or api_key == "":
            raise ValueError(
                "[오류] GEMINI_API_KEY를 찾을 수 없습니다.\n"
                "       로컬 환경: .env 파일을 열어 실제 API 키를 입력했는지 확인하세요.\n"
                "       웹 배포 환경: Streamlit Secrets(Advanced settings)에 "
                "GEMINI_API_KEY를 설정했는지 확인하세요."
            )

        # 2. _knowledge 폴더 확인 (없으면 생성, 비어있으면 경고만)
        self.knowledge_dir.mkdir(parents=True, exist_ok=True)

        # 3. Gemini API 클라이언트 설정 (google-genai 신규 SDK)
        self.client = genai.Client(api_key=api_key)
        self.model_name = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()

        self.knowledge_cache = None  # 캐시 (성능 최적화)

        # 4. 법제처 Open API 클라이언트 (선택 기능)
        #    .env에 LAW_API_OC 값이 없으면 법제처 연동 없이 기존처럼 동작함.
        self.law_client = None
        law_oc_key = os.getenv("LAW_API_OC", "").strip()
        if law_oc_key and _LAW_API_AVAILABLE:
            try:
                self.law_client = LawAPIClient(oc_key=law_oc_key)
            except ValueError as e:
                print(f"[경고] 법제처 API 클라이언트 초기화 실패: {e}")
                print("       법제처 연동 없이 _knowledge 데이터만으로 동작합니다.")
        elif law_oc_key and not _LAW_API_AVAILABLE:
            print("[경고] LAW_API_OC가 설정되어 있지만 law_api.py 파일을 찾을 수 없습니다.")
            print("       법제처 연동 없이 _knowledge 데이터만으로 동작합니다.")

        # 5. 구글 스프레드시트 로깅 (선택 기능)
        #    .env에 GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_PATH가 없으면 로깅 없이 동작함.
        self.sheet_logger = None
        if _SHEET_LOGGER_AVAILABLE:
            self.sheet_logger = SheetLogger()
            if not self.sheet_logger.enabled and (
                os.getenv("GOOGLE_SHEET_ID", "").strip()
                or os.getenv("GOOGLE_CREDENTIALS_PATH", "").strip()
                or os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
            ):
                # 설정을 시도했지만 실패한 경우에만 안내 (아예 설정 안 한 경우는 조용히 비활성)
                print(f"[안내] 구글 시트 로깅이 비활성 상태입니다: {self.sheet_logger.error_message}")

    # ------------------------------------------------------------------
    # Knowledge Base 로드
    # ------------------------------------------------------------------
    def load_knowledge_base(self, force_reload: bool = False) -> str:
        """
        _knowledge 폴더 안의 모든 텍스트/엑셀 자원을 모아 컨텍스트로 변환

        Parameters
        ----------
        force_reload : bool
            True면 캐시를 무시하고 다시 읽음 (지식 베이스 수정 후 즉시 반영하고 싶을 때 사용)

        Returns
        -------
        str
            문서별로 정리된 지식 베이스 텍스트
        """
        if self.knowledge_cache is not None and not force_reload:
            return self.knowledge_cache

        context_chunks = []
        files = sorted(self.knowledge_dir.glob("*.*"))

        if not files:
            print(
                f"[안내] _knowledge 폴더에 파일이 없습니다: {self.knowledge_dir}\n"
                f"       지식 베이스 없이 동작하며, 모든 질의에 '확인 불가' 폴백이 출력됩니다."
            )
            self.knowledge_cache = "(현재 _knowledge 폴더에 등록된 내부 지침 데이터가 없습니다.)"
            return self.knowledge_cache

        load_errors = []

        for file_path in files:
            suffix = file_path.suffix.lower()
            try:
                if suffix in self.SUPPORTED_TEXT_EXT:
                    content = file_path.read_text(encoding="utf-8")
                    if content.strip():
                        context_chunks.append(f"[문서: {file_path.name}]\n{content}\n")

                elif suffix in self.SUPPORTED_EXCEL_EXT:
                    excel_data = pd.read_excel(file_path)
                    if not excel_data.empty:
                        markdown_table = excel_data.to_markdown(index=False)
                        context_chunks.append(f"[문서: {file_path.name}]\n{markdown_table}\n")

                # 그 외 확장자는 조용히 무시 (예: .gitkeep 등)

            except UnicodeDecodeError:
                load_errors.append(f"{file_path.name} - 인코딩 오류 (UTF-8로 저장되었는지 확인 필요)")
            except Exception as e:
                load_errors.append(f"{file_path.name} - {str(e)}")

        if load_errors:
            print("[경고] 일부 지식 베이스 파일을 로드하지 못했습니다:")
            for err in load_errors:
                print(f"   - {err}")

        if not context_chunks:
            print(
                "[안내] _knowledge 폴더의 파일들은 있지만 아직 작성된 내용이 없습니다.\n"
                "       지식 베이스 없이 동작하며, 모든 질의에 '확인 불가' 폴백이 출력됩니다."
            )
            self.knowledge_cache = "(현재 _knowledge 폴더에 등록된 내부 지침 데이터가 없습니다.)"
            return self.knowledge_cache

        self.knowledge_cache = "\n\n---\n\n".join(context_chunks)
        return self.knowledge_cache

    # ------------------------------------------------------------------
    # 법제처 Open API 연동 (선택 기능)
    # ------------------------------------------------------------------
    def fetch_law_context(self, user_question: str, max_articles: int = 5) -> str:
        """
        질문 내용에서 법령명을 추정하여 법제처 API로 관련 조문을 조회

        - law_client가 설정되지 않은 경우(.env에 LAW_API_OC 없음) 빈 문자열 반환
        - API 호출이 실패해도 예외를 던지지 않고 빈 문자열 반환
          (법제처 연동 실패가 전체 답변 생성을 막아서는 안 됨)

        Parameters
        ----------
        user_question : str
            사용자 질문
        max_articles : int
            법령 1건당 가져올 최대 조문 수

        Returns
        -------
        str
            조회된 조문 텍스트. 매칭되는 법령이 없거나 조회 실패 시 빈 문자열.
        """
        if self.law_client is None:
            return ""

        matched_laws = [law for law in self.COMMON_TAX_LAWS if law in user_question]
        if not matched_laws:
            return ""

        law_texts = []
        for law_name in matched_laws[:2]:  # 한 번에 너무 많은 법령을 조회하지 않도록 제한
            try:
                text = self.law_client.get_law_text_by_name(law_name, max_articles=max_articles)
                if text:
                    law_texts.append(text)
            except LawAPIError as e:
                print(f"[경고] 법제처 API 조회 실패 ({law_name}): {e}")
                # 실패해도 계속 진행 (다른 법령 조회는 시도)
                continue

        return "\n\n---\n\n".join(law_texts)

    # ------------------------------------------------------------------
    def _extract_search_keywords(self, user_question: str, max_keywords: int = 3) -> list:
        """
        자연어 질문에서 법제처 검색에 적합한 핵심 키워드 여러 개를 추출.
        Gemini를 가벼운 호출로 한 번 더 사용함. 실패 시 빈 리스트 반환.

        중요 (2026-06-24 실측 확인): 법제처 ntsCgmExpc 검색은 한 번의 호출에
        여러 단어를 쉼표/공백으로 합친 구문을 넣으면 0건이 나오는 경우가 많음
        (API 자체에 OR 연산자가 없고 단순 문자열 매칭이기 때문).
        따라서 "OR 검색"의 효과를 내기 위해, 키워드를 여러 개 추출한 뒤
        각각 별도로 검색 API를 호출하고 결과를 모으는 방식을 사용함
        (이는 search_nts_interpretations에서 수행).
        """
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=(
                    "다음 세무 질문에서, 국세청 법령해석(질의회신) 데이터베이스를 검색할 때 쓸 "
                    f"핵심 단어/용어를 최대 {max_keywords}개까지 추출하세요.\n"
                    "규칙:\n"
                    "- 각 키워드는 띄어쓰기 없는 단일 용어 또는 짧은 복합명사여야 합니다 "
                    "(예: '의제매입세액', '완전포괄주의', '간이과세자').\n"
                    "- 한 줄에 키워드 하나씩만 쓰세요. 한 줄 안에 여러 단어를 띄어쓰기로 나열하지 마세요.\n"
                    "- 일반적인 표현 대신 정확한 법률/세무 전문 용어를 우선하되, "
                    "서로 다른 각도의 키워드를 다양하게 뽑으세요 "
                    "(예: 거래유형, 적용 법령명, 정확한 제도명 등 관점을 다르게).\n"
                    "- 설명, 번호, 불릿 없이 키워드만 한 줄에 하나씩 출력하세요.\n\n"
                    f"질문: {user_question}\n\n"
                    "키워드 목록:"
                ),
            )
            raw_text = (response.text or "").strip() if response.candidates else ""
        except Exception as e:
            print(f"[경고] 검색 키워드 추출 실패: {e}")
            return []

        keywords = []
        for line in raw_text.split("\n"):
            # 모델이 번호/불릿/쉼표를 붙였을 경우를 대비해 정리
            cleaned = line.strip().lstrip("-*0123456789.").strip()
            cleaned = cleaned.split(",")[0].strip()  # 한 줄에 쉼표로 나열했어도 첫 항목만
            if cleaned and cleaned not in keywords:
                keywords.append(cleaned)
            if len(keywords) >= max_keywords:
                break

        return keywords

    # ------------------------------------------------------------------
    # 국세청 법령해석(질의회신/예규) 검색 - 선택 기능 (v1.3 추가)
    # ------------------------------------------------------------------
    def search_nts_interpretations(self, user_question: str, max_results: int = 3) -> str:
        """
        국세청 법령해석(질의회신/예규) 목록을 검색하고, 검색된 안건에 대해
        Gemini의 Google Search grounding 기능으로 본문 내용을 보강 시도.

        중요한 한계 (정확히 이해하고 사용할 것):
        - 법제처 API는 국세청 법령해석에 대해 "목록"만 제공하며, 본문 조회 API는
          존재하지 않음 (다른 부처와 달리 국세청/재정경제부만 본문 API 미제공).
        - 본문은 Google 검색으로 "찾을 수 있으면" 가져오는 것이며, 이는 공식
          데이터가 아니라 검색 결과 기반 참고 정보임. 100% 정확성이 보장되지 않음.
        - 이 기능은 .env의 ENABLE_NTS_SEARCH=true 로 명시적으로 켜야만 작동함.

        Returns
        -------
        str
            안건명/안건번호/해석일자/링크 + (찾은 경우) 검색 기반 본문 요약.
            반드시 "검증 필요한 참고 정보"임을 명시하는 라벨이 포함됨.
            관련 안건이 없거나 기능이 꺼져 있으면 빈 문자열 반환.
        """
        if not self.law_client:
            return ""

        if os.getenv("ENABLE_NTS_SEARCH", "").strip().lower() not in ("true", "1", "yes"):
            return ""

        # 0단계: 자연어 질문에서 핵심 검색 키워드 여러 개 추출
        # (법제처 API는 OR 연산자가 없는 단순 문자열 매칭이므로, 한 번에 여러 단어를
        #  합쳐서 보내면 0건이 나옴. 대신 키워드별로 각각 검색을 호출해 결과를 모으는
        #  방식으로 OR 검색과 동일한 효과를 냄)
        search_keywords = self._extract_search_keywords(user_question, max_keywords=3)
        if not search_keywords:
            return ""

        # 1단계: 각 키워드로 개별 검색 후 결과를 모두 합침 (중복 제거)
        all_cases = []
        seen_ids = set()
        tried_keywords = []
        for kw in search_keywords:
            try:
                kw_cases = self.law_client.search_nts_interpretations(kw, display=max_results)
            except LawAPIError as e:
                print(f"[경고] 국세청 법령해석 검색 실패 (키워드 '{kw}'): {e}")
                continue

            tried_keywords.append(f"{kw}({len(kw_cases)}건)")
            for c in kw_cases:
                if c["id"] not in seen_ids:
                    seen_ids.add(c["id"])
                    all_cases.append(c)

        print(f"[안내] 국세청 법령해석 검색: {', '.join(tried_keywords)} → 중복제거 후 총 {len(all_cases)}건")

        if not all_cases:
            return ""

        # 결과가 많으면 최종적으로 max_results*2 건까지만 사용 (Gemini 검색보강 비용 통제)
        cases = all_cases[: max_results * 2]

        # 2단계: 검색된 안건들로 Gemini Google Search grounding 호출하여 본문 보강 시도
        case_list_text = "\n".join(
            f"- 안건명: {c['title']} / 안건번호: {c['case_no']} / 해석일자: {c['date']} / 링크: {c['link']}"
            for c in cases
        )

        search_prompt = f"""다음은 국세청 법령해석(질의회신/예규) 목록에서 키워드 검색으로 찾은 관련 안건들입니다.
먼저 원래 질문과 가장 관련성이 높은 안건들을 추리고, 그 안건들의 실제 회신 내용(결론)을
Google 검색으로 찾아서, 찾은 내용만 정리해주세요.

[규칙]
1. 원래 질문과 관련성이 낮은 안건은 제외하고, 관련성 높은 안건 위주로 정리하세요.
2. 검색으로 실제 내용을 확인한 안건만 정리하세요.
3. 검색해도 내용을 찾지 못한 안건은 "본문 확인 불가 - 안건명/링크만 제공됨"이라고 명시하세요.
4. 절대로 검색 결과에 없는 내용을 추측하거나 지어내지 마세요.
5. 안건별로 안건번호를 반드시 함께 표시하세요.
6. 이모지는 사용하지 마세요.

[관련 안건 목록]
{case_list_text}

[원래 질문 - 참고용]
{user_question}
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=search_prompt,
                config=genai_types.GenerateContentConfig(
                    tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
                ),
            )
            search_result_text = response.text if response.candidates else ""
        except Exception as e:
            print(f"[경고] Google Search grounding 호출 실패: {e}")
            search_result_text = ""

        if not search_result_text:
            # 검색 보강에 실패해도 목록 자체는 참고 정보로 제공
            return (
                "[국세청 법령해석 - 검색 기반 참고 정보, 본문 미확인]\n"
                "(아래 안건은 관련성이 있을 수 있으나 본문 내용은 확인하지 못했습니다. "
                "회계사가 직접 링크를 확인해야 합니다)\n"
                + case_list_text
            )

        return (
            "[국세청 법령해석 - 검색 기반 참고 정보, 검증 필요]\n"
            "(아래 내용은 Google 검색으로 찾은 참고 정보이며, 법제처/국세청의 공식 본문 API가 아닙니다. "
            "내용의 정확성이 보장되지 않으므로 반드시 링크를 통해 원문을 직접 대조 확인해야 합니다)\n\n"
            + search_result_text
        )

    # ------------------------------------------------------------------
    # 질의 응답 생성
    # ------------------------------------------------------------------
    def generate_guideline(self, user_question: str, thread_history: list = None) -> str:
        """
        사용자 질문과 지식 베이스를 결합하여 Gemini API 호출

        Parameters
        ----------
        user_question : str
            기장 직원의 세무 질의
        thread_history : list[dict], optional
            같은 대화 묶음(스레드) 안의 이전 질의응답 목록.
            각 항목은 {"question": str, "answer": str} 형태.
            제공되면 AI가 이전 질문/답변을 실제로 참고하여 꼬리질문에 맥락 있게 답변함.
            None이거나 빈 리스트면 단발성 질문으로 처리됨(기존 동작과 동일).

        Returns
        -------
        str
            구조화된 마크다운 형식의 회신
        """
        if not user_question or not user_question.strip():
            return "질문이 비어 있습니다. 세무 질의 내용을 입력해주세요."

        knowledge_text = self.load_knowledge_base()
        law_context = self.fetch_law_context(user_question)
        nts_context = self.search_nts_interpretations(user_question)

        # ------------------------------------------------------------------
        # 대화 흐름 블록 ("지금 무슨 대화를 하고 있는가" — 지식베이스와는 완전히 별개 레이어)
        #
        # 설계 의도 (2026-06-25 수정):
        # - 지식베이스는 "참고 자료 창고"이고, thread_history는 "지금 사용자와 나눈 대화 메모"임.
        #   둘은 성격이 다르므로, 지식베이스 분량이 커도 대화 흐름이 묻히면 안 됨.
        # - 기존에는 history_block을 지식베이스 바로 뒤, 사용자 질문 바로 앞에 살짝 끼워 넣는
        #   방식이었음. 지식베이스가 길어지면(수만 자) 모델이 바로 앞의 지식베이스 내용에
        #   끌려가 꼬리질문의 맥락(예: "쉽게 설명해줘")을 직전 대화가 아닌 지식베이스의
        #   다른 주제로 잘못 연결하는 사례가 실측됨 (예: 원천세 질문 후 "쉽게 설명해줘" →
        #   의제매입세액 답변이 나오는 오작동).
        # - 해결: "지금 진행 중인 대화"와 "새 질문"을 하나의 묶음으로 묶어 프롬프트 맨 앞에
        #   배치하고, 역할을 "최우선 판단 기준"으로 명시. 지식베이스는 그 뒤에 "참고 자료"로
        #   배치하여 역할을 분명히 구분함.
        # ------------------------------------------------------------------
        if thread_history:
            history_lines = []
            for i, turn in enumerate(thread_history, start=1):
                history_lines.append(f"[이전 질의 {i}] {turn['question']}\n[이전 회신 {i}]\n{turn['answer']}")

            conversation_block = (
                "[지금 진행 중인 대화 — 최우선 판단 기준]\n"
                "아래는 같은 사용자가 이번 대화 묶음 안에서 지금까지 주고받은 질문과 답변입니다. "
                "사용자가 '새 주제 시작'을 누르지 않은 이상, 지금의 새 질문은 원칙적으로 "
                "이 대화의 후속/꼬리질문입니다(예: '더 쉽게 설명해줘', '그래서 결론이 뭐야', "
                "'그거 말고 다른 경우는?' 등). 새 질문만 보면 의미가 불완전하더라도, "
                "반드시 직전 질문/답변의 주제와 흐름을 이어서 해석하세요. "
                "아래쪽에 나오는 [지식베이스 참고 자료]는 세법 근거를 찾기 위한 배경 자료일 뿐이며, "
                "지금 대화의 주제를 바꾸거나 다른 주제로 끌고 가는 근거로 쓰면 안 됩니다.\n\n"
                + "\n\n".join(history_lines)
            )

            question_block = (
                f"{conversation_block}\n\n"
                f"[지금 사용자의 새 질문 — 위 대화의 후속질문으로 해석할 것]\n{user_question}"
            )
        else:
            question_block = f"[사용자 질문 — 이번 대화의 첫 질문]\n{user_question}"

        context_blocks = []
        if law_context:
            context_blocks.append(f"[법제처 조문 데이터 — 참고 자료]\n{law_context}")
        if nts_context:
            context_blocks.append(
                f"[국세청 법령해석 검색 참고정보 — 참고 자료, 검증 필요]\n{nts_context}"
            )
        extra_context = ("\n\n" + "\n\n".join(context_blocks)) if context_blocks else ""

        full_prompt = f"""{question_block}

위 질문에 답변하기 위해, 아래의 참고 자료를 근거 우선순위에 따라 활용하세요.
(주의: 아래 자료는 세법 근거를 찾기 위한 배경 자료입니다. 위에서 파악한 '지금 대화의 주제'를
바꾸는 용도로 사용하지 마세요.)

[지식베이스 참고 자료 — 1순위 근거]
{knowledge_text}{extra_context}

위의 [회신 양식]을 반드시 준수하여 답변하세요."""

        response = self.client.models.generate_content(
            model=self.model_name,
            contents=full_prompt,
            config=genai_types.GenerateContentConfig(
                system_instruction=SYSTEM_INSTRUCTION,
                # temperature 낮춤 (2026-06-25 추가): 세무 자문은 창작이 아니라 정확성이
                # 생명인 도메인이므로, 기본값(보통 1.0 근처)보다 훨씬 낮춰서 AI가 화려한
                # 표현이나 임의의 추정을 줄이고, 가진 근거(지식베이스/법령) 기반으로만
                # 정형화된 답변을 내도록 함. 0.1~0.3 범위가 일반적으로 권장되며,
                # 너무 0에 가깝게 하면 가끔 어색하게 반복되는 문장이 나올 수 있어 0.2로 설정.
                temperature=0.2,
            ),
        )

        if not response.candidates:
            return "[오류] AI가 응답을 생성하지 못했습니다. 질문을 다시 입력하거나 잠시 후 시도해주세요."

        return response.text

    def generate_guideline_with_retry(self, user_question: str, max_retries: int = 3, thread_history: list = None) -> str:
        """
        API 호출 재시도 로직 포함 (지수 백오프)
        """
        last_error = None
        for attempt in range(max_retries):
            try:
                return self.generate_guideline(user_question, thread_history=thread_history)
            except Exception as e:
                last_error = e
                if attempt < max_retries - 1:
                    wait_seconds = 2 ** attempt
                    print(
                        f"[경고] API 호출 실패 (시도 {attempt + 1}/{max_retries}). "
                        f"{wait_seconds}초 후 재시도합니다. 원인: {e}"
                    )
                    time.sleep(wait_seconds)

        return (
            "[오류] API 호출에 반복적으로 실패했습니다. "
            f"잠시 후 다시 시도해주세요.\n오류 상세: {last_error}"
        )

    # ------------------------------------------------------------------
    # 결과 저장
    # ------------------------------------------------------------------
    def save_response(self, question: str, response: str, output_dir: Path) -> str:
        """
        질의 및 회신을 마크다운 파일로 저장

        Parameters
        ----------
        question : str
            사용자 질문
        response : str
            AI 회신
        output_dir : Path
            저장할 디렉토리

        Returns
        -------
        str
            저장된 파일 경로
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"자문_{timestamp}.md"
        filepath = output_dir / filename

        content = f"""# 세무질의 자문 기록

**질의일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 질의 내용
{question}

---

## AI 회신

{response}

---
*포스원 회계법인 세무질의 AI 시스템 자동 생성*
"""
        filepath.write_text(content, encoding="utf-8")
        return str(filepath)

    # ------------------------------------------------------------------
    # 확정 PIN 관리 - v1.5 추가
    # ------------------------------------------------------------------
    def _pin_file_path(self) -> Path:
        base_dir = Path(__file__).resolve().parent
        return base_dir / ".confirm_pin.txt"

    def has_pin_set(self) -> bool:
        """PIN이 설정되어 있는지 확인"""
        return self._pin_file_path().exists()

    def set_pin(self, new_pin: str) -> None:
        """
        확정 PIN을 설정/변경 (해시값만 저장, 평문 저장하지 않음)

        Parameters
        ----------
        new_pin : str
            새로 설정할 PIN (숫자/문자 조합, 4자 이상 권장)
        """
        import hashlib
        pin_hash = hashlib.sha256(new_pin.strip().encode("utf-8")).hexdigest()
        self._pin_file_path().write_text(pin_hash, encoding="utf-8")

    def verify_pin(self, input_pin: str) -> bool:
        """입력된 PIN이 설정된 PIN과 일치하는지 확인"""
        import hashlib
        pin_path = self._pin_file_path()
        if not pin_path.exists():
            return False
        stored_hash = pin_path.read_text(encoding="utf-8").strip()
        input_hash = hashlib.sha256(input_pin.strip().encode("utf-8")).hexdigest()
        return stored_hash == input_hash

    # ------------------------------------------------------------------
    # 지식베이스 확정 저장 - v1.5 추가
    # ------------------------------------------------------------------
    def confirm_to_knowledge_base(
        self,
        question: str,
        confirmed_content: str,
        target_file: str = "01_공통_세무질의회신집.txt",
    ) -> str:
        """
        회계사가 검증/확정한 질의응답을 _knowledge 폴더의 파일에 추가 저장

        주의: 이 메서드를 호출하기 전에 반드시 verify_pin()으로 PIN을 검증해야 합니다.
        이 메서드 자체는 PIN 검증을 하지 않습니다 (호출 측 책임).

        Parameters
        ----------
        question : str
            원래 질문
        confirmed_content : str
            확정할 내용 (회계사가 검토/수정한 최종 텍스트)
        target_file : str
            추가할 _knowledge 폴더 내 파일명 (기본: 공통 세무질의회신집)

        Returns
        -------
        str
            저장된 파일의 전체 경로
        """
        target_path = self.knowledge_dir / target_file
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        entry = (
            f"\n\n질의: {question}\n"
            f"회신요지: {confirmed_content}\n"
            f"근거: (회계사 확정 — {timestamp})\n"
            f"비고: 포스원 세무 자문 AI 시스템을 통해 확정됨\n"
            f"---\n"
        )

        with open(target_path, "a", encoding="utf-8") as f:
            f.write(entry)

        # 캐시 무효화 - 다음 질의부터 즉시 반영되도록
        self.knowledge_cache = None

        return str(target_path)


# ----------------------------------------------------------------------
# 터미널 단독 실행 테스트
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("포스원 회계법인 세무 자문 AI 엔진 - 터미널 테스트")
    print("=" * 60)

    try:
        engine = TaxAdvisorEngine()
        print(f"[정상] 엔진 초기화 완료. Knowledge Base 경로: {engine.knowledge_dir}\n")
    except (FileNotFoundError, ValueError) as e:
        print(str(e))
        sys.exit(1)

    try:
        kb_text = engine.load_knowledge_base()
        print(f"[정상] 지식 베이스 로드 완료 ({len(kb_text)}자)\n")
    except ValueError as e:
        print(str(e))
        sys.exit(1)

    print("질문을 입력하세요 (종료: 빈 줄 입력 후 Enter)\n")

    while True:
        question = input("질문> ").strip()
        if not question:
            print("\n종료합니다.")
            break

        print("\n조회 중입니다 (3~5초 소요)...\n")
        answer = engine.generate_guideline_with_retry(question)
        print("-" * 60)
        print(answer)
        print("-" * 60)

        save = input("\n이 결과를 저장하시겠습니까? (y/n): ").strip().lower()
        if save == "y":
            output_dir = os.getenv("OUTPUT_DIR", "").strip()
            if not output_dir:
                output_dir = str(Path(__file__).resolve().parent / "세법검토_아카이브")
            saved_path = engine.save_response(question, answer, Path(output_dir))
            print(f"\n저장 완료: {saved_path}\n")
        print()
