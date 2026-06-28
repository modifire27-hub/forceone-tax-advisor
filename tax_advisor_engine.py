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

### 3. 실무 회계처리 지침 (해당하는 경우만)
[이 섹션 전체는, 질문에 분개를 일으킬 수 있는 "구체적인 거래"(예: 특정 비용 지출,
특정 매출 발생, 특정 자산 취득 등 실제 전표 처리가 필요한 사건)가 있을 때만
작성하세요.

질문이 "제도가 무엇인지", "기준이 어떻게 되는지", "의무가 언제 발생하는지" 같은
일반적인 제도/개념 설명을 묻는 경우(구체적인 거래 당사자나 금액이 특정되지
않은 질문)에는, 이 섹션 전체(추천 차변/대변 계정과목, 표준 적요 가이드, 세무조정
항목 모두)를 생략하세요. 이런 질문에 회계처리를 억지로 끼워 넣으면, 그 거래를
"내가 지출하는 입장"인지 "내가 매출을 발생시키는 입장"인지 등 방향을 알 수 없는
상태에서 추측으로 채우게 되어, 차변/대변이 거꾸로 되는 등의 오류로 이어집니다.
예를 들어 "현금영수증 의무발행 제도가 무엇인가요"처럼 일반 제도 설명 질문에는
회계처리 섹션을 생략하고, "음식점에서 직원 회식비 10만원을 카드로 결제했는데
이 비용을 어떻게 처리하나요"처럼 구체적 거래가 있는 질문에는 작성하세요.

작성하는 경우, 그 거래의 당사자가 "발급/수취" 중 어느 쪽인지, 비용 지출인지
매출 발생인지를 질문의 사실관계에서 정확히 판단한 뒤 방향을 정하세요. 특히
증빙(세금계산서, 현금영수증, 계산서 등) 관련 질문은 "내가 발급하는 입장
(매출 발생 → 차변 현금/외상매출금, 대변 매출+부가세예수금)"과 "내가 받는 입장
(매입/비용 지출 → 차변 비용/자산+부가세대급금, 대변 현금/미지급금)"을 혼동하기
쉬우니, 질문에 명시된 사실관계를 다시 확인해 방향을 정하세요.]
- **추천 차변 계정과목**: [과목명만. 코드는 표기하지 마세요 — 회사마다 계정코드 체계가 다를 수 있습니다]
- **추천 대변 계정과목**: [과목명만]
- **표준 적요 가이드**: [기장 직원이 전표에 입력할 표준 적요 텍스트 예시. 실제
  거래 연월을 알 수 없으므로, 연월 표기는 영문 대문자([YYYY], [MM]) 대신
  "○○○○년 ○○월"이라는 자리표시자를 사용하세요. 이 자리표시자는
  "연도 부분(○○○○년)"과 "월 부분(○○월)"이 합쳐진 것으로, 적요 문장 안에서
  연월을 나타낼 자리에 정확히 한 번만 그대로 옮겨 적으세요. 절대로 "○○월"을
  추가로 한 번 더 적어 연월 표기가 두 번 겹치게 만들지 마세요. 차량번호 등
  다른 가변 정보도 "○○○○" 형태로 해당 자리에 한 번만 표기하세요.
  예: "○○○○년 ○○월 차량유지비 지급 (차량번호 ○○○○)"]
- **세무조정이 필요한 경우** (해당하는 경우만): 손금불산입/손금산입 등 세무조정이
  필요하다면, 어떤 항목을 어떤 절차로 처리해야 하는지 일반적인 흐름(예: "세무조정
  명세서에 해당 금액을 손금불산입으로 반영하고, 소득처분을 통해 귀속자에게
  상여/배당 등으로 처분") 위주로 설명하세요. 특정 회계 프로그램의 구체적인 메뉴명이나
  메뉴 경로는 언급하지 마세요 — 프로그램마다 메뉴 구성이 다르고, 정확히 검증되지 않은
  메뉴명을 안내하면 오히려 실무자를 혼란스럽게 할 위험이 있습니다.
  해당사항이 없으면 이 항목은 생략하세요.

### 4. 관련 근거 및 유의사항
- **법령 근거**: [참고한 법령명과 조항. 없으면 생략]
- **예규/판례 참고**: [참고한 안건명, 안건번호. 검색 기반 참고정보라면 원문 대조 필요 명시. 없으면 생략]
- **일반 실무 기준**: [법령/예규 근거 없이 일반 세무 실무 지식으로 답한 부분이 있다면 명시. 없으면 생략]
- **실무 유의사항**: [기장 직원이 놓치기 쉬운 세무적 리스크 및 점검사항]
"""


def _extract_marked_section(text: str, start_marker: str, end_markers: list) -> str:
    """
    구분자([검증결과], [수정요약] 등) 사이의 텍스트를 안전하게 추출하는 헬퍼.
    모델이 형식을 어겨 구분자가 없으면 빈 문자열을 반환함 (호출하는 쪽에서
    안전한 기본값으로 폴백 처리).
    """
    if start_marker not in text:
        return ""
    after = text.split(start_marker, 1)[1]
    cut_positions = [after.find(m) for m in end_markers if m in after]
    if cut_positions:
        after = after[: min(cut_positions)]
    return after.strip()


def _extract_mismatch_items(verification_text: str) -> list:
    """
    검증 결과 텍스트에서 "답변과 다름"으로 판정된 항목만 골라내는 헬퍼.

    배경 (D안 — 2026-06-26 추가, 검증 기능 고도화 3차 보강):
    - B안(출처 링크), C안(볼드 표시)을 거쳤지만, 실사용 테스트에서 더 근본적인
      문제가 반복 확인됨: 검증(1차 호출)이 "답변과 다름"을 정확히 찾아내고
      정확한 대체 근거까지 제시해도, 수정(2차 호출)이 그 항목들 중 일부를
      누락한 채 "전체를 반영했다"고 잘못 보고하는 사례가 두 번 연속 발생함.
    - 원인으로 추정되는 것: _apply_corrections가 검증 결과 텍스트 "전체"를
      통째로 모델에게 주고 "답변과 다름 항목을 알아서 찾아서 고쳐라"라고
      맡기는 구조였음. 검증 결과가 길어질수록(항목이 10개, 20개로 늘어날수록)
      모델이 그 안에서 정확히 어떤 항목들이 "답변과 다름"인지 스스로 다시
      찾아내는 과정에서 일부를 놓치는 것으로 보임.
    - 해결: "답변과 다름" 항목을 모델에게 다시 찾으라고 시키지 않고, 코드가
      직접 텍스트를 파싱해서 미리 추출함. 이렇게 추출된 목록을
      _apply_corrections의 프롬프트 맨 앞에 "반드시 반영해야 할 N개 항목"으로
      별도로, 작고 명확하게 제시함 — 모델의 일은 "이 정해진 목록을 빠짐없이
      반영하는 것"으로 좁아지고, "전체 텍스트에서 답변과 다름을 찾아내는 것"
      이라는 부담이 사라짐.
    - 추가로, _apply_corrections가 작업을 끝낸 뒤 이 추출된 항목들이 실제로
      반영됐는지 사후 점검하는 안전장치(_check_unresolved_mismatches)에도
      이 함수가 재사용됨.

    파싱 규칙:
    - 검증 결과는 "- (항목) → (판정)" 형태의 불릿 목록으로 옴. 각 항목은
      여러 줄에 걸쳐 들여쓰기로 이어질 수 있음(예: 예시에 나온 멀티라인 불릿).
    - 따라서 "- "로 시작하는 줄을 새 항목의 시작으로 보고, 다음 "- "가
      나오기 전까지(또는 텍스트 끝까지)를 하나의 항목으로 묶음.
    - 그 항목 텍스트 안에 "답변과 다름"이라는 문자열이 있으면 그 항목
      전체(여러 줄 포함)를 결과 리스트에 추가함.
    - 모델이 형식을 다소 어겨도(예: 불릿 기호가 "-" 대신 "*"인 경우 등)
      최대한 견고하게 동작하도록, 줄 시작이 "-" 또는 "*"인 경우를 모두
      새 항목의 시작으로 인정함. 그래도 못 찾으면 빈 리스트를 반환하고,
      호출하는 쪽은 "추출 실패"를 "답변과 다름이 0개"와는 다르게 처리해야
      안전함 (아래 _apply_corrections에서 처리).

    Parameters
    ----------
    verification_text : str
        _run_verification_search가 반환한 검증 결과 전체 텍스트

    Returns
    -------
    list[str]
        "답변과 다름"이 포함된 항목 텍스트(여러 줄 포함)의 리스트.
        형식이 전혀 인식되지 않으면(불릿 자체를 못 찾으면) 빈 리스트.
    """
    if not verification_text:
        return []

    lines = verification_text.split("\n")
    items = []
    current_item_lines = []

    def _flush():
        if current_item_lines:
            items.append("\n".join(current_item_lines).strip())

    for line in lines:
        stripped = line.strip()
        is_new_bullet = stripped.startswith("- ") or stripped.startswith("* ")
        if is_new_bullet:
            _flush()
            current_item_lines = [line]
        elif current_item_lines:
            # 직전 불릿 항목이 여러 줄로 이어지는 경우 (들여쓰기된 후속 줄)
            current_item_lines.append(line)
        # current_item_lines가 비어있는데 불릿도 아닌 줄(머리말 등)은 무시
    _flush()

    mismatch_items = [item for item in items if "답변과 다름" in item]
    return mismatch_items


def _check_unreflected_items(original_content: str, corrected_content: str, mismatch_items: list) -> list:
    """
    사후 안전장치 — _apply_corrections가 항목별 치환을 마친 뒤, 그래도
    각 mismatch 항목에서 지적된 "원본의 틀린 핵심 문구"가 최종본에 여전히
    남아있는지 코드로 한 번 더 직접 대조함.

    역할 변경 (E안 — 2026-06-26): D안 시기에는 이 함수가 "모델이 전체
    재작성을 했는데 실제로 반영됐는지"를 확인하는 유일한 검증 수단이었음.
    E안으로 구조가 바뀌면서 _apply_corrections는 이제 항목별로 직접
    str.replace 치환을 수행하므로, 치환에 사용한 원문 발췌가 더 이상
    corrected_content에 없는 것이 기본적으로 보장됨(치환이 성공했다면
    원문 그대로는 사라지고 대체 문장으로 바뀌어 있어야 함). 따라서 이
    함수는 "주된 검증 수단"이 아니라, 혹시 모델이 치환 대상과 약간 다른
    표현으로 같은 오류를 본문 다른 곳에 중복 기술해 둔 경우를 잡아내는
    이중 점검(2차 안전망) 용도로만 사용됨.

    설계 의도:
    - 모델이 "전부 반영했습니다"라고 자체 보고해도, 실제로는 일부가 그대로
      남아있는 사례가 반복 확인됐음(예: "수정 사항: 제78조를 제42조로
      수정했습니다"라고 보고했지만 최종본에 제78조가 그대로 남아있던 사례).
      이런 자기모순을 모델의 보고만으로는 잡을 수 없으므로, 코드가 직접
      "원본에서 틀렸다고 지적된 표현이 최종본에도 똑같이 있는가"를 단순
      문자열 대조로 확인함.
    - 완벽한 검증은 아님(모델이 표현을 살짝 바꿔서 같은 오류를 다른 말로
      반복할 수도 있고, 반대로 우연히 일부 단어가 겹쳐서 오탐이 날 수도
      있음). 그래서 이 함수의 결과는 "확정 판정"이 아니라 회계사에게
      보여줄 경고용 참고 정보로만 사용함 — 저장 자체를 막지는 않음.

    파싱 규칙:
    - 각 mismatch 항목 텍스트에서, 원본의 틀린 표현은 보통 다음 두 가지
      패턴 중 하나로 인용됨(검증 프롬프트의 출력 형식 지시에 따름):
        (a) 화살표 앞부분: "원본 조항/문구 자체" (예: "부가가치세법 시행령
            제78조 제1항 제3호 → 답변과 다름...")
        (b) 본문 중 **"..."** 형태로 볼드 처리된 인용구 (예: 원본은
            **"연체료는 손금불산입될 수 있다"**고만 했으나...)
    - (a)는 화살표(→) 앞쪽 텍스트를 그대로 검사 문구로 사용함. 단, 이 앞쪽
      텍스트가 "법령 조항 번호"처럼 구체적인 경우에만 의미가 있으므로,
      너무 짧은 문구(2글자 미만)는 오탐 위험이 커서 건너뜀.
    - (b)는 정규식으로 **"..."** 안의 내용을 모두 추출해 검사 문구로 사용함.
    - 둘 중 하나라도 corrected_content 안에 등장하면(즉 안 고쳐졌다고 의심되면),
      이 항목을 "반영 안 됨 의심"으로 분류함.

    Parameters
    ----------
    original_content : str
        원본 답변 (참고용, 현재는 직접 사용하지 않지만 추후 확장을 위해 유지)
    corrected_content : str
        모델이 반환한 "수정된전체본"
    mismatch_items : list[str]
        _extract_mismatch_items로 추출된 "답변과 다름" 항목 리스트

    Returns
    -------
    list[str]
        반영이 안 된 것으로 의심되는 원본 문구(짧게 잘라서) 목록.
        의심되는 게 없으면 빈 리스트.
    """
    import re

    if not corrected_content:
        return []

    unresolved_snippets = []

    for item in mismatch_items:
        suspect_phrases = []

        # (a) 화살표 앞부분 (원본 조항/항목명)
        if "→" in item:
            before_arrow = item.split("→", 1)[0]
            before_arrow = before_arrow.lstrip("-*").strip()
            if len(before_arrow) >= 6:  # 너무 짧으면 오탐 위험이 커서 건너뜀
                suspect_phrases.append(before_arrow)

        # (b) 볼드 인용구 **"..."** 안의 내용
        quoted = re.findall(r'\*\*["“]([^"”*]{6,})["”]\*\*', item)
        suspect_phrases.extend(quoted)

        for phrase in suspect_phrases:
            # 검사 문구 자체가 corrected_content에 그대로 남아있으면 반영 안 됐다고 의심
            if phrase in corrected_content:
                snippet = phrase[:40] + ("..." if len(phrase) > 40 else "")
                if snippet not in unresolved_snippets:
                    unresolved_snippets.append(snippet)
                break  # 이 항목은 이미 의심 처리했으니 다음 항목으로

    return unresolved_snippets


def _find_snippet_in_content(snippet: str, content: str):
    """
    F안 — 다단계 매칭 (2026-06-26 추가, E안의 치환 성공률 보강).

    배경:
    - E안은 모델이 발췌한 original_snippet이 content 안에 "한 글자도 안
      틀리고" 그대로 있어야만 치환에 성공하는 구조였음. 실사용에서, 모델이
      검증 항목 텍스트에서 발췌를 재인용할 때 공백이나 문장부호를 미묘하게
      다르게 옮기는 경우가 있어, 검증은 정확했는데 치환만 실패하는 사례가
      발생함(예: "기타 비용 700만원 초과분..." 항목).
    - 이 함수는 완전 일치가 안 되더라도 점점 느슨한 기준으로 재시도해서,
      "치환 가능한 위치를 더 적극적으로 찾아보는" 역할을 함. 다만 너무
      느슨하게 풀면 엉뚱한 곳을 잘못 치환할 위험이 있으므로, 단계를 명확히
      나누고 각 단계가 안전한 범위 내에서만 동작하도록 제한함.

    매칭 단계 (위에서부터 순서대로 시도, 처음 성공하는 단계를 사용):
    1. "exact" — snippet이 content에 완전히 그대로 존재. 가장 안전함.
    2. "normalized" — 공백(스페이스, 탭, 줄바꿈)의 개수/종류 차이를 무시하고
       비교. 일치하면, content 안에서 정규화 기준으로 매칭된 실제 원문
       구간(공백 포함 원래 형태)을 찾아 그 구간을 반환함 — 치환 시 원본의
       줄바꿈 등 형태를 그대로 보존하기 위함.
    3. "partial" — snippet의 앞부분(약 20자, 너무 짧으면 건너뜀)이 content
       안에서 발견되면, 그 시작 위치부터 다음 마침표(. 또는 다음 줄바꿈)
       까지를 "실제 교체할 문장 단위"로 잡아 반환함. snippet 전체가 약간
       다르게 재인용됐어도, 적어도 시작 지점은 맞는 경우가 많다는 점을
       활용함. 너무 짧은 접두부로 매칭하면 오탐 위험이 있으므로, 접두부가
       15자 미만이면 이 단계는 건너뜀.
    위 세 단계 모두 실패하면 None을 반환함 (호출하는 쪽은 "자동 수정
    실패"로 정직하게 보고해야 함 — 억지로 매칭시키지 않음).

    Parameters
    ----------
    snippet : str
        모델이 발췌했다고 주장하는 원본 텍스트
    content : str
        실제로 검색할 대상 텍스트(원본 답변 또는 치환 진행 중인 중간 결과물)

    Returns
    -------
    tuple[str, str] | None
        (실제로 content 안에서 찾아낸, 치환 대상이 될 정확한 부분 문자열, 매칭 단계 이름)
        못 찾으면 None.
    """
    import re

    if not snippet or not content:
        return None

    # 1단계: 완전 일치
    if snippet in content:
        return (snippet, "exact")

    # 2단계: 공백 정규화 일치
    # snippet과 content 양쪽의 연속 공백(스페이스/탭/줄바꿈)을 단일 스페이스로
    # 치환한 뒤 비교. content 쪽은 원래 위치 정보를 보존하기 위해, 정규화된
    # 패턴을 정규식으로 만들어 원본 content에서 직접 검색함.
    normalized_snippet = re.sub(r"\s+", " ", snippet).strip()
    if normalized_snippet and len(normalized_snippet) >= 6:
        # snippet의 각 "단어"를 이어주는 부분에 \s+ 가 들어가도 매칭되도록
        # 정규식을 구성 (단어 자체는 그대로, 단어 사이 공백만 유연하게 처리)
        words = normalized_snippet.split(" ")
        pattern = r"\s+".join(re.escape(w) for w in words if w)
        if pattern:
            try:
                match = re.search(pattern, content)
            except re.error:
                match = None
            if match:
                return (match.group(0), "normalized")

    # 3단계: 접두부 부분 일치 — snippet 앞부분으로 시작 위치를 찾고,
    # 그 지점부터 다음 문장 끝(마침표 또는 줄바꿈)까지를 교체 단위로 삼음.
    prefix_len = min(20, len(normalized_snippet))
    if prefix_len >= 15:
        prefix = normalized_snippet[:prefix_len]
        # content에서 prefix와 가장 유사한 시작 지점을 찾기 위해, prefix의
        # 공백도 유연하게 처리한 정규식으로 검색
        prefix_words = prefix.split(" ")
        prefix_pattern = r"\s+".join(re.escape(w) for w in prefix_words if w)
        if prefix_pattern:
            try:
                match = re.search(prefix_pattern, content)
            except re.error:
                match = None
            if match:
                start = match.start()
                # 시작 지점부터, 다음 마침표(.) 또는 줄바꿈을 만날 때까지를
                # 교체 대상 구간으로 삼음 (문장 단위 교체로, 의미 단위가
                # 깨지지 않도록 함). 마침표/줄바꿈을 못 찾으면 snippet과
                # 비슷한 길이만큼만 자름 (전체 문서를 통째로 날리는 사고 방지).
                remainder = content[start:]
                end_match = re.search(r"[.\n]", remainder)
                if end_match:
                    end = start + end_match.end()
                else:
                    end = start + min(len(snippet) + 20, len(remainder))
                matched_text = content[start:end]
                if matched_text:
                    return (matched_text, "partial")

    return None


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
        _knowledge 폴더 안의 모든 텍스트/엑셀 자원 + 구글시트 '지식베이스' 탭 내용을
        모아서 하나의 컨텍스트로 변환.

        설계 의도 (2026-06-25 추가 — 웹/PC 동기화):
        - 로컬 _knowledge 폴더의 .txt/.xlsx 파일은 "최초 기본 지식"으로, PC에서 직접
          수정한 뒤 GitHub에 다시 올려야만 갱신됨 (드물게 일괄 정리할 때만 바뀜).
        - 회계사가 "지식베이스에 확정 저장"으로 새로 추가하는 지식은 이제 로컬
          파일이 아니라 구글시트의 '지식베이스' 탭에 쌓임 (PC/웹 어디서 저장해도
          같은 시트에 들어가므로 항상 동기화됨).
        - 따라서 AI에게 줄 전체 지식베이스 = 로컬 파일 내용 + 구글시트 지식베이스
          탭 내용. 구글시트 연동이 비활성 상태(.env/Secrets 미설정)라면 로컬 파일
          내용만으로도 정상 동작함 (하위 호환, 필수 기능 아님).

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

        # 구글시트 '지식베이스' 탭 내용 추가 (연동 비활성이면 빈 문자열이 와서 자동 생략됨)
        if self.sheet_logger and self.sheet_logger.enabled:
            sheet_knowledge_text = self.sheet_logger.get_all_knowledge_text()
            if sheet_knowledge_text:
                context_chunks.append(
                    f"[문서: 구글시트_지식베이스_확정내용]\n{sheet_knowledge_text}\n"
                )

        if not context_chunks:
            print(
                "[안내] 등록된 지식 베이스 데이터가 없습니다 (로컬 파일도, 구글시트 지식베이스도 비어 있음).\n"
                "       지식 베이스 없이 동작하며, 모든 질의에 '확인 불가' 폴백이 출력됩니다."
            )
            self.knowledge_cache = "(현재 등록된 내부 지침 데이터가 없습니다.)"
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
    # 지식베이스 확정 저장 전 자동 검증 - v1.6 추가
    # ------------------------------------------------------------------
    # 설계 의도 (2026-06-25 추가):
    # - 회계사가 "지식베이스에 확정 저장"을 누르기 전에, AI 답변 내용을 다시 한번
    #   웹검색으로 교차 검증하고, 어느 _knowledge 파일에 저장하면 좋을지도 자동으로
    #   추천함. 회계사가 매번 "이거 맞나? 어느 파일에 넣지?"를 직접 판단하는 부담을
    #   줄이고, AI가 1차로 점검한 결과를 보고 최종 승인만 내리도록 함.
    # - 검증은 100% 확신을 주는 게 아니라 "의심되는 부분이 있는지"를 알려주는
    #   참고용임을 명확히 해야 함. 검증 결과가 "문제없음"이라고 나와도 회계사의
    #   최종 판단(PIN 승인)은 여전히 필요함.
    KNOWLEDGE_FILE_OPTIONS = [
        "01_공통_세무질의회신집.txt",
        "02_업종별_기장유의사항.txt",
        "04_부가세_처리지침.txt",
        "05_기타_세법_예규.txt",
    ]

    def verify_before_confirm(
        self,
        question: str,
        content: str,
        precomputed_verification_text: str = None,
    ) -> dict:
        """
        지식베이스 확정 저장 전, 답변 내용을 웹검색으로 재검증하고
        저장할 파일을 추천함.

        Parameters
        ----------
        question : str
            원래 질문
        content : str
            확정 저장하려는 답변 내용 (AI 답변 전체 또는 회계사가 수정한 버전)
        precomputed_verification_text : str, optional
            (v1.6 추가 — 교차검증 스레드 기능) 이미 검증이 끝난 텍스트가 있으면
            여기로 전달함. 전달되면 이 메서드는 _run_verification_search를
            다시 호출하지 않고(웹검색 중복 실행 방지), 곧바로 ②③ 단계
            (_apply_corrections, _recommend_knowledge_file)만 수행함.
            None이면(기본값) 기존 동작과 완전히 동일하게 이 메서드가 직접
            웹검색부터 새로 수행함 — 기존 호출부(버튼 한 번으로 검증+수정을
            한꺼번에 끝내는 기존 화면)는 이 파라미터를 그냥 생략하면 한 글자도
            바뀔 일 없이 그대로 동작함.

        Returns
        -------
        dict
            {
                "verification_text": str,  # 검증 결과 설명 (의심 포인트, 근거 등)
                "corrected_content": str,  # 검증 결과를 반영해 다시 작성한 최종본
                "correction_summary": str, # 어디를 왜 고쳤는지 짧은 요약 (1~3줄)
                "recommended_file": str,   # 추천 저장 파일명
                "recommended_reason": str, # 추천 이유 (한 줄)
            }
            웹검색 호출 자체가 실패하면 verification_text에 실패 사유가 담기고,
            corrected_content는 원본(content)을 그대로 반환함 (검증/수정 실패가
            저장 자체를 막지는 않음 — 최종 판단은 항상 회계사의 PIN 승인).

        설계 의도 (2026-06-25 — 검증과 수정을 분리된 두 번의 AI 호출로 변경):
        - 1차 시도: 검증과 수정을 한 번의 호출로 같이 시켰음(검증결과+수정본을 한
          응답 안에 같이 작성). 결과: 검증 단계에서는 오류를 정확히 찾아내면서도,
          같은 응답의 "수정된최종본" 섹션에는 그 오류가 그대로 남아있는 자기 모순이
          반복됨.
        - 2차 시도: 출력 순서를 "수정된최종본을 먼저, 검증결과를 나중에"로 바꿔서
          모델이 먼저 고친 글을 쓰고 그걸 설명하게 함. 결과: 똑같은 자기 모순이
          다시 발생함 — 즉 "한 번의 응답 안에서 일관성을 유지하라"는 지시 자체가
          신뢰할 수 없는 전제였음이 확인됨.
        - 3차 — 최종 구조: 검증과 수정을 완전히 별개의 두 번의 모델 호출로 분리함.
          1차 호출(_run_verification_search)은 원본을 검토해 "어디가 왜 틀렸는지"
          오류 목록만 만들고, 텍스트 수정은 전혀 하지 않음. 2차 호출
          (_apply_corrections)은 그 오류 목록과 원본 텍스트를 둘 다 명시적인
          입력으로 받아, "이 목록에 적힌 대로 원본을 처음부터 다시 써라"는 단순하고
          명확한 단일 작업만 수행함. 이렇게 하면 2차 호출의 모델이 "직전에 내가
          무엇을 생각했는지" 기억해 이어갈 필요가 없어, 검증과 수정 사이의 불일치
          위험이 구조적으로 줄어듦.

        설계 의도 (v1.6 추가 — 2026-06-28, 교차검증 스레드 기능):
        - 실사용 회고(7.2.1절)에서 "검색 결과를 사람이 먼저 확인한 뒤 그 결과를
          가지고 AI에게 재질문하는 방식이 훨씬 정확하다"는 핵심 인사이트가
          확인됨. 이를 반영해, 확정 저장 전 단계에 "회계사가 검증 결과를 보고
          다른 AI에게 교차 확인을 요청하고, 그 답변을 다시 시스템에 꼬리질문
          으로 입력해 재검토를 반복할 수 있는" 별도 작업공간(검증 스레드)을
          신설함 — render_confirm_to_kb_workspace의 새 단계들 참고.
        - 이 검증 스레드는 자체적으로 여러 차례 웹검색 기반 재검증
          (run_cross_check_verification)을 반복하며 "최종 검증 텍스트"를
          만들어냄. 회계사가 "확정 단계로 진행"을 누르면, 그 최종 검증
          텍스트를 이 메서드에 precomputed_verification_text로 넘겨서
          웹검색을 한 번 더 새로 하지 않고 곧바로 수정 단계로 넘어가게 함.
        """
        if precomputed_verification_text is not None:
            verification_text = precomputed_verification_text
        else:
            verification_text = self._run_verification_search(question, content)

        corrected_content, correction_summary = self._apply_corrections(
            question, content, verification_text
        )
        recommended_file, recommended_reason = self._recommend_knowledge_file(
            question, corrected_content
        )

        return {
            "verification_text": verification_text,
            "corrected_content": corrected_content,
            "correction_summary": correction_summary,
            "recommended_file": recommended_file,
            "recommended_reason": recommended_reason,
        }

    def run_initial_verification(self, question: str, content: str) -> str:
        """
        교차검증 스레드의 1라운드 진입점 (v1.6 신규, 2026-06-28).

        기존 _run_verification_search(언더스코어 메서드 — 클래스 내부에서만
        쓰일 것을 가정해 설계됨)를 화면(streamlit_ui.py)에서 직접 호출하기
        위한 public 래퍼. _run_verification_search 자체는 한 글자도 바꾸지
        않고 그대로 재사용함 — 기존 verify_before_confirm 흐름에 영향 없음.
        """
        return self._run_verification_search(question, content)

    def _run_verification_search(self, question: str, content: str) -> str:
        """
        1차 호출 — 검증 전용. 원본 답변을 검토해 "어디가 왜 틀렸는지" 목록만
        만들고, 텍스트 수정은 전혀 하지 않음 (수정은 _apply_corrections가 별도로 함).

        변경 (C안 — 2026-06-26, 검증 기능 고도화 2차 시도):
        - B안(검색 출처 링크를 검증 결과에 첨부)은 실사용 테스트 결과 폐기함.
          grounding_chunks가 law.go.kr 같은 신뢰 가능한 출처와 youtube.com,
          naver.com, tistory.com 같은 무관한 출처가 뒤섞여 나오고, 검증 항목과
          출처 사이의 1:1 대응 관계도 없어 회계사가 "어느 링크가 어느 항목의
          근거인지" 알 수 없었음. 결과적으로 클릭해서 대조할 동기 자체가
          생기지 않는 출처 목록이 되어 실무 가치가 낮았음.
        - C안으로 전환: 검증 결과에서 "답변과 다름"으로 판정된 항목을 볼드
          마크다운(**...**)으로 강조하도록 모델에 명시 지시함. 회계사가 검증
          상세 내용을 펼쳤을 때, 굵게 표시된 부분만 눈으로 훑으면 어디가
          틀렸는지 바로 파악할 수 있게 하는 것이 목적. 이 볼드 표시는
          _apply_corrections의 최종 수정본에도 동일하게 이어져, "검증에서
          틀렸다고 한 부분"과 "실제로 고쳐진 부분"을 회계사가 빠르게 대조할
          수 있게 함.

        변경 (2026-06-26 추가 보강 — 실사용 사례에서 발견된 문제 수정):
        - 실사용 테스트에서, 검증 단계가 "원본의 시행령 조항이 틀렸다"는 것은
          정확히 찾아냈지만, 정확한 대체 근거는 "일반적인 해석입니다" 같은
          모호한 표현으로 회피하고 끝낸 사례가 발생함. 그 결과 2차 호출
          (_apply_corrections)이 받은 검증 결과에 "무엇으로 고쳐야 하는지"가
          없어서, 틀린 조항 번호가 최종본에 그대로 남는 문제가 생김 (검증은
          맞았는데 수정이 반영 안 되는 자기모순의 새로운 유형).
        - 이를 막기 위해 3단계 지침에 "조항이 틀렸다는 사실만 적지 말고, 정확한
          대체 근거(정확한 조항 번호 또는 법리 구조)까지 검색해서 찾아내라"는
          의무를 추가함. "일반적인 해석"처럼 모호한 표현으로 마무리하는 것을
          명시적으로 금지하고, 충분히 검색을 반복한 뒤에만 "확인하지 못했습니다"
          라고 쓸 수 있도록 4단계 지침도 함께 강화함.
        """
        verify_prompt = f"""다음은 세무 자문 AI가 생성한 답변입니다. 이 답변을 그대로 회계법인의
내부 지식베이스(향후 다른 질문에도 참고 자료로 쓰일 데이터)에 확정 저장하기 전에,
Google 검색으로 이 답변의 핵심 주장들이 실제 법령/실무와 맞는지 점검해주세요.

이번 단계에서는 검증만 하면 됩니다. 답변 텍스트를 고치거나 다시 쓰지 마세요 —
그 작업은 다음 단계에서 별도로 진행됩니다. 지금은 "무엇이 틀렸는지"를 정확하고
빠짐없이 찾아내는 데에만 집중하세요.

[원래 질문]
{question}

[검증할 답변 내용]
{content}

[작업 순서 — 반드시 순서대로, 빠짐없이 수행하세요]

1단계 — 검증 대상 목록화:
답변 안에서 검증이 필요한 모든 항목을 빠짐없이 나열하세요. 다음 종류를 모두
포함해야 합니다 — 일부만 골라서 검증하면 안 됩니다:
  (a) 법령/시행령/시행규칙의 조항 번호 전부 (본문에 나온 모든 "제○조", "제○항")
  (b) 구체적인 수치(비율 %, 한도액, 기간, 세율 등) 전부
  (c) "~로 처리됩니다", "~됩니다", "~해야 합니다" 같은 단정적 결론 문장 전부
  (d) 사례에 등장하는 여러 하위 케이스가 서로 다른 결론을 가질 수 있는 경우
      (예: "보험 종류별로", "사업자 유형별로", "거래 형태별로" 등 분류가 있다면,
      그 분류의 각 항목이 정말로 답변에서 구분되어 다뤄졌는지 확인 — 답변이
      여러 하위 케이스를 하나의 결론으로 뭉뚱그렸다면 그것도 의심 항목입니다)

2단계 — 각 항목 개별 검색:
1단계에서 나열한 항목을 하나씩, 따로따로 Google 검색하세요. 여러 항목을 한 번에
묶어서 검색하면 정확도가 떨어지므로 항목당 최소 1회씩 검색하세요. "전반적으로
맞는 것 같다"는 인상만으로 판단하지 말고, 각 항목을 실제로 검색해 확인한
결과만을 근거로 삼으세요.

3단계 — 불일치 정리:
검색 결과와 답변 내용이 다른 항목을 모두 찾아, 정확히 어떻게 다른지(원본이 뭐라고
했는지, 검색 결과로는 무엇이 맞는지) 구체적으로 정리하세요. 다음 단계에서 이
정리 내용을 바탕으로 원본을 다시 쓸 것이므로, 모호하게 적지 말고 "다음 사람이
그대로 받아서 고칠 수 있을 만큼" 구체적으로 적으세요.

[중요 — 조항/법령 근거가 틀린 경우의 추가 의무] 원본이 인용한 법령/시행령
조항 번호가 검색 결과와 다르다는 것을 확인했다면, "이 조항이 틀렸다"는 사실만
적고 끝내지 마세요. 반드시 한 단계 더 검색해서, "그러면 정확한 근거는 무엇인가
(정확한 조항 번호, 또는 정확한 법리 구조)"까지 찾아내세요. 예를 들어 원본이
인용한 특정 조항이 실제로는 전혀 다른 내용을 규정하고 있다는 것을 발견했다면,
"그 결론을 실제로 뒷받침하는 근거가 무엇인지"(다른 조항, 또는 두 조항이 결합되어
적용되는 구조 등)를 추가로 검색해서 다음 사람이 그대로 받아서 원본의 틀린 조항을
대체할 수 있을 만큼 구체적으로 적으세요. "일반적인 해석", "~인 것으로 보임" 같은
모호한 표현으로 마무리하지 말고, 검색으로 실제 확인된 정확한 근거를 명시하세요.

4단계 — 검증되지 않은 부분 처리:
추측하지 말고, 실제로 검색해서 확인한 내용만 근거로 판단하세요. 다만 위 3단계의
의무를 다하기 위해 검색을 충분히 반복했는데도 정확한 대체 근거를 끝내 찾지
못한 경우에만 "이 부분은 검색으로 명확히 확인하지 못했습니다"라고 솔직하게
밝히세요 — 한두 번 검색해보고 바로 이 표현으로 넘어가지 마세요. "일치함"
또는 "특이사항 없음"은 실제로 검색해서 확인했을 때만 쓰세요 — 검색을 안 했거나
결과가 불충분해서 판단을 못 한 경우에는 절대 그렇게 쓰면 안 됩니다.

[출력 형식 — 마크다운 볼드 표시 규칙 반드시 준수]
1단계에서 나열한 항목을 모두, 하나도 빠짐없이 다음 형식으로 나열하세요:
- (검증 대상 항목) → (검색 결과: 일치함 / 답변과 다름 — 구체적으로 어떻게 다른지 / 확인 못함)

다음 두 가지는 반드시 마크다운 볼드(**텍스트**)로 감싸서 강조하세요. 이는
회계사가 검증 결과를 빠르게 훑어볼 때, 굵은 글씨만 보고 어디가 문제인지
바로 파악하도록 돕기 위한 표시입니다:
  - "답변과 다름"이라는 판정 자체: **답변과 다름**
  - 그 항목에서 실제로 잘못된 부분(원본의 어떤 표현이 잘못됐는지)의 핵심 문구:
    예) 원본은 **"연체료는 손금불산입될 수 있다"**고만 했으나
"일치함"이나 "확인 못함" 판정에는 볼드를 사용하지 마세요 — 볼드는 오직
"답변과 다름"으로 판정된 항목에서만 사용해, 회계사의 시선이 진짜 문제
지점에만 집중되게 하세요.

예시:
- 법인세법 시행령 제43조 제2항 → 일치함
- 건강보험 연체료 손금불산입 여부 → **답변과 다름**. 원본은 **"연체료는
  손금불산입될 수 있다"**고만 했으나, 검색 결과 건강보험은 연체금·가산금
  모두 손금불산입, 국민연금·고용보험·산재보험은 가산금만 손금불산입이고
  연체금은 손금산입됨. 원본에 이 보험 종류별 구분이 전혀 없음.
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=verify_prompt,
                config=genai_types.GenerateContentConfig(
                    tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
                    temperature=0.2,
                ),
            )
            verification_text = response.text if response.candidates else ""
        except Exception as e:
            verification_text = f"[검증 실패] 웹검색 기반 재검증 호출에 실패했습니다: {e}"

        if not verification_text:
            verification_text = "[검증 실패] 검증 결과를 받지 못했습니다. 회계사가 직접 내용을 확인해주세요."
        return verification_text

    # ------------------------------------------------------------------
    # 교차검증 스레드 — v1.6 신규 (2026-06-28)
    # ------------------------------------------------------------------
    # 배경: 7.2.1절 회고에서 확인된 핵심 인사이트 — "검색 결과를 사람이 먼저
    # 확인한 뒤 그 결과를 가지고 AI에게 재질문하는 방식이 훨씬 정확하다"를
    # 실제 기능으로 구현한 것. 기존의 "AI 단독 자동 검증+수정"(verify_before_
    # confirm 한 번 호출로 끝) 흐름은 그대로 유지하면서, 그 앞 단계에 회계사가
    # 원하는 만큼 반복할 수 있는 "외부 AI 교차검증" 라운드를 추가함.
    #
    # 신규질문이든 검색기록(구글시트)에서 불러온 과거 질문이든, 이 교차검증
    # 스레드는 항상 동일한 형태(question + answer로 시작하는 독립된 메모리상
    # thread_history)로 다뤄짐 — UI 쪽(streamlit_ui.py)의
    # render_confirm_to_kb_workspace가 이 thread_history를
    # session_state[f"{key_prefix}_verification_thread"]로 관리함.
    def build_cross_check_prompt(
        self,
        question: str,
        content: str,
        verification_text: str,
        verification_thread_history: list = None,
    ) -> str:
        """
        회계사가 다른 AI(ChatGPT, Gemini 웹, Perplexity 등)에게 그대로 복사해서
        붙여넣을 수 있는 교차검증 요청 문구를 만듦.

        버그 수정 (2026-06-28 — 회계사 피드백 반영):
        - 기존 버전은 라운드와 무관하게 문구 안에 "1차 검증한 결과"라는 말이
          고정되어 있었음. verification_text 자체는 매 라운드 최신 내용으로
          정상 갱신되고 있었지만, 그 내용을 감싸는 안내문이 항상 "1차"라고
          박혀 있어서, 회계사가 보기에 "과거 문장이 그대로 나온다"는 인상을
          주는 문제가 있었음.
        - 2차 이상의 라운드에서는, 그 이전까지 외부 AI가 지적했던 내용과
          그걸 반영해 갱신된 검증 결과까지 함께 보여줘야, 다른 AI가 이미
          다뤄진 사항을 또 처음부터 반복하지 않고 새로운 부분에 집중할 수
          있음. verification_thread_history를 받아 이를 포함시킴.

        Parameters
        ----------
        question : str
            원래 질문
        content : str
            검증 대상 원본 답변 (또는 그 시점까지의 최신 수정본)
        verification_text : str
            가장 최근 검증 결과 텍스트 (1차든, N차 재검증 결과든 상관없이 항상
            "지금 시점의 최신" 텍스트를 그대로 넘기면 됨)
        verification_thread_history : list[dict], optional
            이전 라운드 기록. 비어있거나 None이면 1차 검증으로 간주함.

        Returns
        -------
        str
            다른 AI에게 그대로 붙여넣을 수 있는 완성된 질문 문구
        """
        round_no = (len(verification_thread_history) if verification_thread_history else 0) + 1

        if verification_thread_history:
            prior_blocks = []
            for r in verification_thread_history:
                prior_blocks.append(
                    f"[{r['round']}차 검증 결과]\n{r['verification_text']}\n\n"
                    f"[{r['round']}차 검증 후 받았던 다른 AI 답변]\n"
                    f"{r.get('external_ai_input', '(없음)')}"
                )
            prior_history_text = (
                "\n\n[이전 라운드 진행 기록 — 이미 다뤄진 사항이니 참고만 하고, "
                "되도록 새로 찾은 부분 위주로 답해주세요]\n" + "\n\n---\n\n".join(prior_blocks)
            )
            latest_label = f"{round_no}차 검증 결과 — 직전 라운드까지의 지적을 반영해 갱신된 최신 결과입니다"
        else:
            prior_history_text = ""
            latest_label = "1차 검증 결과"

        return f"""아래는 세무 자문 AI 시스템이 생성한 답변과, 그 답변을 Google 검색으로
검증한 결과입니다. 이 내용을 교차검증해주세요.

[원래 질문]
{question}

[AI가 생성한 답변]
{content}
{prior_history_text}

[{latest_label} — 이미 의심되는 부분이 표시되어 있습니다]
{verification_text}

[요청 사항]
1. 위 검증 결과에서 "답변과 다름"으로 표시된 항목들이 실제로 맞는지 직접
   검색하거나 알고 있는 지식으로 다시 확인해주세요.
2. 지금까지의 검증이 놓쳤을 수 있는 다른 오류(법령 조항 번호, 수치, 단정적
   결론 등)도 추가로 점검해주세요.
3. 결론적으로 이 답변에서 실제로 고쳐야 할 부분이 무엇인지, 그리고 정확한
   근거(조항 번호, 정확한 수치 등)는 무엇인지 구체적으로 알려주세요.
4. 확실하지 않은 부분은 "확실하지 않음"이라고 솔직하게 말해주세요. 추측해서
   단정적으로 답하지 마세요.
"""

    def run_cross_check_verification(
        self,
        question: str,
        content: str,
        verification_thread_history: list,
        external_ai_input: str,
    ) -> str:
        """
        회계사가 외부 AI로부터 받아온 답변을 반영해, 웹검색을 동반한 재검증을
        한 번 더 수행함 (교차검증 스레드의 2회차 이상 라운드).

        설계 의도:
        - _run_verification_search와 마찬가지로 Google Search 도구를 그대로
          사용함 — "다른 AI 답변만 보고 판단"이 아니라, 그 답변이 맞는지도
          시스템이 직접 한 번 더 검색해서 교차 확인하는 것이 목표.
        - 이번 라운드까지의 검증 스레드 전체 기록(verification_thread_history)을
          함께 제공하여, 이전 라운드에서 이미 확인된 사항을 또 처음부터
          반복해서 검색하지 않고, "새로 들어온 외부 AI 답변" 위주로 점검하게
          유도함.
        - 반환값은 새로운 "최종 검증 텍스트"이며, 형식은 _run_verification_search
          결과와 동일한 "- (항목) → (판정)" 불릿 목록 구조를 유지하도록 지시함.
          이렇게 해야 이후 verify_before_confirm에 precomputed_verification_text로
          넘겼을 때 _extract_mismatch_items가 정상적으로 항목을 파싱할 수 있음
          (기존 _apply_corrections 로직을 그대로 재사용하기 위한 호환성 유지).

        Parameters
        ----------
        question : str
            원래 질문
        content : str
            검증 대상 원본 답변 (최초 답변 — 라운드가 반복되어도 바뀌지 않음)
        verification_thread_history : list[dict]
            이번 라운드 이전까지의 검증 스레드 기록.
            각 항목은 {"round": int, "verification_text": str, "external_ai_input": str} 형태.
        external_ai_input : str
            회계사가 이번에 새로 붙여넣은 외부 AI의 답변 텍스트

        Returns
        -------
        str
            이번 라운드까지 반영된 새로운 검증 결과 텍스트 (기존 _run_verification_search와
            동일한 불릿 형식)
        """
        prior_rounds_text = ""
        if verification_thread_history:
            round_blocks = []
            for r in verification_thread_history:
                round_blocks.append(
                    f"[{r['round']}차 검증 결과]\n{r['verification_text']}\n\n"
                    f"[{r['round']}차 검증 후 회계사가 외부 AI에게 받아온 답변]\n"
                    f"{r.get('external_ai_input', '(없음)')}"
                )
            prior_rounds_text = (
                "\n\n[이전까지의 검증 진행 기록 — 참고만 하고, 이미 다뤄진 항목을 "
                "또 처음부터 반복 검색하지 마세요]\n" + "\n\n---\n\n".join(round_blocks)
            )

        verify_prompt = f"""다음은 세무 자문 AI가 생성한 답변에 대해, 회계사가 다른 AI에게
교차검증을 요청해서 받아온 답변입니다. 이 새 답변을 참고하여, Google 검색으로
한 번 더 직접 확인한 뒤, 최종 검증 결과를 정리해주세요.

[원래 질문]
{question}

[검증 대상 원본 답변]
{content}
{prior_rounds_text}

[이번에 회계사가 외부 AI로부터 받아온 답변]
{external_ai_input}

[작업]
1. 외부 AI가 지적한 사항이 실제로 맞는지, Google 검색으로 직접 다시 확인하세요.
   외부 AI의 말을 그대로 믿지 말고, 검색으로 확인된 사실만 최종 근거로
   삼으세요 — 외부 AI도 틀릴 수 있습니다.
2. 외부 AI가 새로 지적한 사항과, 이전 라운드에서 이미 확인된 사항을 모두
   합쳐서, "원본 답변과 비교했을 때 최종적으로 무엇이 다른지"를 다시
   정리하세요. 이미 일치한다고 확인된 항목은 다시 검색할 필요 없습니다.
3. 검색으로도, 외부 AI의 답변으로도 명확히 확인되지 않는 부분은 "확인 못함"
   이라고 솔직하게 남기세요.

[출력 형식 — 기존과 동일한 형식을 반드시 유지하세요]
검증이 필요한 모든 항목을 다음 형식으로 나열하세요(이전 라운드에서 이미
"일치함"으로 확정된 항목도 빠짐없이 포함하세요 — 누락되면 다음 단계에서
그 항목이 사라진 것으로 처리됩니다):
- (검증 대상 항목) → (검색 결과: 일치함 / 답변과 다름 — 구체적으로 어떻게 다른지 / 확인 못함)

"답변과 다름"이라는 판정과, 그 항목에서 실제로 잘못된 부분(원본의 어떤 표현이
잘못됐는지)의 핵심 문구는 마크다운 볼드(**텍스트**)로 감싸서 강조하세요.
"일치함"이나 "확인 못함" 판정에는 볼드를 사용하지 마세요.
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=verify_prompt,
                config=genai_types.GenerateContentConfig(
                    tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
                    temperature=0.2,
                ),
            )
            new_verification_text = response.text if response.candidates else ""
        except Exception as e:
            new_verification_text = f"[재검증 실패] 교차검증 재검토 호출에 실패했습니다: {e}"

        if not new_verification_text:
            new_verification_text = (
                "[재검증 실패] 재검증 결과를 받지 못했습니다. 회계사가 직접 내용을 확인해주세요."
            )
        return new_verification_text

    def _apply_corrections(self, question: str, content: str, verification_text: str) -> tuple:
        """
        2차 호출 — 수정 전용. 1차 호출의 검증 결과(오류 목록)와 원본 텍스트를
        받아 원본을 고침. 검증 자체는 다시 하지 않음(이미 1차에서 끝남).

        변경 (E안 — 2026-06-26 추가, 검증 기능 고도화 4차 — 구조 전환):
        - B/C/D안은 모두 "원본 문서 전체를 모델이 한 번에 다시 써내는" 구조를
          유지한 채 프롬프트나 사후 점검만 보강한 것이었음. 그런데 실사용에서
          "모델이 무엇을 고쳐야 하는지 정확히 알고 있고, 자기가 수정 사항
          요약에 그렇게 적었음에도, 실제 전체 재작성 결과물에는 그 수정이
          빠지는" 사례가 D안의 사후 경고로도 계속 잡혔음 (예: 본문은 고쳤지만
          법령 근거 목록은 깜빡함). 근본 원인은 "문서 전체를 기억하며
          처음부터 다시 만들어내는" 작업 자체가, 분량이 늘어날수록 일부
          내용을 놓치기 쉬운 작업이라는 점.
        - E안은 구조를 바꿈: 모델에게 "문서 전체를 다시 써라"고 시키지 않고,
          "답변과 다름" 항목 하나당 1번씩 별도로 작은 호출
          (_resolve_single_mismatch)을 해서, ① 원본에서 문제가 된 부분의
          정확한 발췌(원문 그대로, 한 글자도 안 바꾼 인용)와 ② 그것을 대체할
          새 문장만 받아옴. 그 다음 코드가 파이썬 문자열 치환(str.replace)으로
          원본의 그 발췌가 등장하는 모든 위치를 새 문장으로 직접 바꿔 끼움.
        - 이렇게 하면 모델의 작업이 "문서 전체를 재현하며 기억"하는 게 아니라
          "이 한 항목에 대한 전/후 문장 쌍을 만드는" 작은 단위로 줄어들고,
          항목들이 서로 영향을 주지 않고 독립적으로 처리됨. 또한 치환은
          코드가 직접 수행하므로 "반영했다고 말했지만 실제로는 안 됨" 자체가
          구조적으로 불가능해짐 — 발췌가 원본에 실제로 존재하면 반드시
          치환되고, 존재하지 않으면(모델이 원문을 잘못 인용한 경우) 그 사실이
          그대로 결과에 기록되어 회계사에게 보고됨.
        - 같은 오류 문구가 본문과 "법령 근거" 목록 등 여러 곳에 반복 등장하는
          경우도, str.replace가 기본적으로 모든 일치 위치를 한 번에 바꾸므로
          "본문만 고치고 근거 목록은 빠뜨리는" 문제가 원천적으로 없어짐.
        - 볼드 강조(C안)는 치환된 새 문장 자체를 코드에서 **로 감싸서 적용함
          (모델에게 볼드를 시키지 않고 코드가 직접 처리 — 더 확실함).

        변경 (F안 — 2026-06-26 추가, 치환 성공률 보강):
        - 실사용에서, 검증(1차 호출)이 정확하게 오류를 잡아냈는데도, 모델이
          그 항목의 발췌(original_snippet)를 원문과 한 글자라도 다르게
          재인용해서 단순 str 포함 검사(in)가 실패하는 사례가 발견됨(예:
          "기타 비용 700만원 초과분..." 항목). 이 경우 실제로는 치환할 수
          있는 위치가 분명히 존재하는데도 "자동 수정 실패"로 보고되어
          회계사가 매번 수동으로 고쳐야 하는 불편이 있었음.
        - _find_snippet_in_content 헬퍼로 매칭을 다단계화함: 완전 일치가
          안 되면 공백 정규화 일치를, 그래도 안 되면 접두부 기반 부분
          일치(문장 단위)를 시도함. 느슨한 매칭으로 치환된 경우는
          correction_summary에 "(유사 매칭으로 치환됨 — 확인 권장)"이라고
          표시해, 완전 일치 치환보다 신뢰도가 낮을 수 있음을 투명하게 알림.
        - 세 단계 모두 실패하는 경우만 기존처럼 "자동 수정 실패"로 정직하게
          보고함 — 무리하게 매칭을 강행해 엉뚱한 부분을 잘못 고치는 것보다,
          못 찾으면 솔직히 못 찾았다고 하는 원칙은 그대로 유지함.

        Returns
        -------
        tuple
            (corrected_content: str, correction_summary: str)
        """
        mismatch_items = _extract_mismatch_items(verification_text)

        if not mismatch_items and len(verification_text) > 50:
            return content, "수정 사항 없음 (검증 결과에 '답변과 다름' 항목이 없어 원본을 그대로 사용합니다)"

        if not mismatch_items:
            # verification_text가 비정상적으로 짧아 파싱 자체가 의미 없는 경우의
            # 폴백 — 이 경우는 안전하게 원본을 그대로 사용함 (E안 구조상
            # "항목이 없는데 전체를 고쳐라"는 호출 자체가 의미가 없으므로,
            # 기존 D안처럼 모델에게 전체를 떠넘기지 않음)
            return content, "검증 결과 형식을 인식하지 못해 자동 수정을 진행하지 못했습니다. 회계사가 직접 검토해주세요."

        corrected_content = content
        applied_summaries = []
        failed_items = []

        for item in mismatch_items:
            resolution = self._resolve_single_mismatch(question, item)
            original_snippet = resolution.get("original_snippet", "").strip()
            replacement = resolution.get("replacement", "").strip()
            short_reason = resolution.get("short_reason", "").strip()

            if not original_snippet or not replacement:
                failed_items.append(
                    f"(발췌/교체문을 생성하지 못함) {item[:80]}..."
                )
                continue

            # F안 — 다단계 매칭 (2026-06-26 추가, 치환 성공률 보강):
            # 모델이 "원문 그대로"를 발췌하라고 지시받아도, 실제로는 공백/
            # 줄바꿈/문장부호를 미묘하게 다르게 재현하는 경우가 실사용에서
            # 확인됨(완전 일치 1건이 실패하면 그 항목 전체가 자동 수정 불가로
            # 보고됨). _find_snippet_in_content가 완전 일치 → 공백 정규화
            # 일치 → 핵심 구절 부분 일치(앞부분으로 시작 위치를 찾고 다음
            # 마침표까지를 교체 단위로 사용) 순서로 점점 느슨하게 시도함.
            match_result = _find_snippet_in_content(original_snippet, corrected_content)

            if match_result is None:
                # 4단계(신규, 2026-06-28 추가 — 회계사 피드백 반영):
                # 1~3단계는 모두 "원문에 실제로 존재하는 발췌를 정확히(또는
                # 거의 정확히) 찾아서 그 부분만 좁게 교체"하는, 창작 없는
                # 가장 안전한 방식임. 그런데 외부 AI 교차검증을 여러 라운드
                # 거치면서 검증 항목의 표현이 원본 문구와 점점 멀어지는
                # 경우, 모델이 만든 original_snippet이 원문 어디에도 못
                # 찾아지는 사례가 늘어나 "자동 수정 실패"가 반복됨.
                # 해결: 1~3단계가 모두 실패한 항목에 대해서만, 이 항목이
                # 본문의 어느 단락을 가리키는지 모델에게 다시 물어 그 단락
                # 전체를 새로 작성해 교체하는 보조 경로를 추가함. 발췌-치환
                # 방식보다 신뢰도는 낮으므로(원본 단락 식별과 재작성을 모델
                # 판단에 맡김), 1~3단계가 모두 실패했을 때만 마지막 수단으로
                # 시도하고, 이 경로로 처리된 항목은 회계사가 한 번 더
                # 확인하도록 요약에 명시함.
                keyword_resolution = self._resolve_mismatch_by_keyword(
                    question, item, corrected_content
                )
                target_paragraph = keyword_resolution.get("target_paragraph", "").strip()
                new_paragraph = keyword_resolution.get("new_paragraph", "").strip()
                kw_reason = keyword_resolution.get("short_reason", "").strip()

                if target_paragraph and new_paragraph and target_paragraph in corrected_content:
                    corrected_content = corrected_content.replace(
                        target_paragraph, f"**{new_paragraph}**", 1
                    )
                    reason_note = (kw_reason or f"\"{original_snippet[:30]}...\" 관련 단락 재작성")
                    reason_note += " (키워드 기반 재작성 — 확인 권장)"
                    applied_summaries.append(reason_note)
                else:
                    failed_items.append(
                        f"(원문에서 정확히 일치하는 위치를 찾지 못함, 키워드 기반 재작성도 실패) "
                        f"\"{original_snippet[:60]}...\""
                    )
                continue

            actual_text_to_replace, match_strategy = match_result
            corrected_content = corrected_content.replace(
                actual_text_to_replace, f"**{replacement}**", 1
            )
            reason_note = short_reason or f"\"{original_snippet[:30]}...\" 수정"
            if match_strategy != "exact":
                # 완전 일치가 아닌 느슨한 매칭으로 치환된 경우, 회계사가 한 번
                # 더 눈으로 확인하면 좋다는 신호를 요약에 남겨둠 (저장을 막지는
                # 않지만, "정확도가 살짝 낮을 수 있다"는 투명성 확보)
                reason_note += " (유사 매칭으로 치환됨 — 확인 권장)"
            applied_summaries.append(reason_note)

        if applied_summaries:
            correction_summary = "; ".join(applied_summaries)
        else:
            correction_summary = "수정 사항 없음"

        if failed_items:
            warning = (
                "[자동 수정 실패 경고] 검증에서 지적된 항목 중 "
                f"{len(failed_items)}건은 시스템이 자동으로 고치지 못했습니다. "
                "회계사가 아래 내용을 직접 확인해 수동으로 고쳐주세요:\n"
                + "\n".join(f"  - {f}" for f in failed_items)
                + "\n\n"
            )
            correction_summary = warning + correction_summary

        # 이중 안전망: 항목별 치환이 끝난 뒤에도, 같은 오류가 본문 다른 곳에
        # 치환 대상과 다른 표현으로 중복 기술되어 남아있지 않은지 한 번 더 점검.
        # (위 failed_items는 "치환을 시도했으나 위치를 못 찾은 경우"만 잡으므로,
        # 이 점검은 그와 별개로 "치환은 성공했지만 다른 곳에 같은 오류가 또
        # 있는 경우"를 잡기 위한 것)
        residual = _check_unreflected_items(content, corrected_content, mismatch_items)
        if residual:
            residual_warning = (
                "[이중 점검 경고] 자동 점검 결과, 위 수정 후에도 일부 오류 표현이 "
                "문서의 다른 위치에 남아있을 수 있습니다 (참고용, 오탐 가능). "
                "확인해주세요:\n"
                + "\n".join(f"  - {snippet}" for snippet in residual)
                + "\n\n"
            )
            correction_summary = residual_warning + correction_summary

        return corrected_content, correction_summary

    def _resolve_single_mismatch(self, question: str, mismatch_item: str) -> dict:
        """
        E안의 핵심 — "답변과 다름" 항목 하나만 떼어서, 모델에게 단 두 가지만
        받아오는 작은 전용 호출:
          1. original_snippet: 원본 문서에서 문제가 된 부분의 정확한 발췌
             (한 글자도 바꾸지 않은 원문 그대로)
          2. replacement: 그 발췌를 대체할 새 문장(검증 결과의 정답 반영)

        이 호출은 문서 전체를 보지 않고 mismatch_item 하나만 입력으로 받음 —
        "문서 전체를 기억하며 재현"하는 부담이 전혀 없으므로, 모델이 다른
        부분에 신경 쓸 필요 없이 이 한 항목에만 집중함.

        검색 도구는 사용하지 않음 — 이미 1차 호출(_run_verification_search)이
        검색을 통해 mismatch_item 안에 정답을 충분히 적어뒀으므로, 이 단계는
        그 내용을 그대로 옮겨 적는 단순 변환 작업임.

        Returns
        -------
        dict
            {
                "original_snippet": str,  # 원본에서 그대로 찾을 발췌. 못 만들면 빈 문자열.
                "replacement": str,       # 대체할 새 문장. 못 만들면 빈 문자열.
                "short_reason": str,      # 1줄 요약 (예: "조항 번호를 제42조로 수정")
            }
            호출 자체가 실패하면 모든 값이 빈 문자열인 dict를 반환함 (호출하는
            쪽에서 "발췌/교체문을 생성하지 못함"으로 처리).
        """
        resolve_prompt = f"""다음은 세무 자문 문서의 검증 결과 중 "답변과 다름"으로 판정된 항목
하나입니다. 이 항목을 보고, 원본 문서에서 고쳐야 할 부분의 정확한 발췌와
그 대체 문장을 만들어주세요.

[원래 질문 — 참고용]
{question}

[검증 결과 항목 — 이 안에 "원본이 뭐라고 했는지"와 "정답이 무엇인지"가
모두 들어있습니다]
{mismatch_item}

[작업]
1. 위 항목 안에서, 원본 문서에 실제로 쓰여 있던 표현을 그대로(단 한 글자도
   바꾸지 않고) 찾아서 "original_snippet"으로 적으세요. 보통 **"..."** 형태로
   인용되어 있거나, 화살표(→) 앞에 조항 번호/항목명으로 나타나 있습니다.
   - 너무 짧게 자르면(예: 조항 번호만) 문서 안에 같은 표현이 여러 군데
     있어서 의도와 다른 곳까지 바뀔 위험이 있으니, 문맥을 식별할 수 있는
     만큼 충분히 길게(가능하면 한 문장 단위로) 잡으세요.
   - 단, 너무 길게 잡아서 검증 항목에 안 나온 내용까지 추측해서 포함하면
     안 됩니다 — 항목에 실제로 언급된 범위 내에서만 발췌하세요.
2. 검증 결과에 적힌 정답 내용을 반영해서, 그 발췌를 대체할 새 문장을
   "replacement"로 작성하세요. 원본과 같은 어조(전문가 답변체)를 유지하고,
   원본 발췌가 문장의 일부였다면 새 문장도 문맥상 자연스럽게 이어지도록
   작성하세요.
3. 무엇을 어떻게 고쳤는지 5~15자 내의 짧은 한국어 구절로 "short_reason"에
   요약하세요. (예: "시행령 제78조를 제42조로 수정")
4. 검증 항목 자체가 "확인 못함"이거나 모호해서 명확한 대체 문장을 만들 수
   없다면, 모든 값을 빈 문자열로 두세요. 억지로 만들어내지 마세요.

[출력 형식 — 다른 설명 없이 이 JSON 형식만 출력하세요. 마크다운 코드블록
표시(```)도 쓰지 마세요]
{{"original_snippet": "...", "replacement": "...", "short_reason": "..."}}
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=resolve_prompt,
                config=genai_types.GenerateContentConfig(temperature=0.1),
            )
            raw_text = response.text.strip() if response.candidates else ""
        except Exception as e:
            print(f"[경고] 항목별 수정 호출 실패: {e}")
            return {"original_snippet": "", "replacement": "", "short_reason": ""}

        if not raw_text:
            return {"original_snippet": "", "replacement": "", "short_reason": ""}

        # 모델이 지시를 어기고 ```json 코드블록으로 감싸는 경우를 대비한 정리
        cleaned = raw_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("```")[1]
            if cleaned.startswith("json"):
                cleaned = cleaned[4:]
            cleaned = cleaned.strip()

        try:
            import json
            parsed = json.loads(cleaned)
        except (json.JSONDecodeError, ValueError) as e:
            print(f"[경고] 항목별 수정 결과 JSON 파싱 실패: {e} / 원본: {raw_text[:200]}")
            return {"original_snippet": "", "replacement": "", "short_reason": ""}

        return {
            "original_snippet": str(parsed.get("original_snippet", "")),
            "replacement": str(parsed.get("replacement", "")),
            "short_reason": str(parsed.get("short_reason", "")),
        }

    def _resolve_mismatch_by_keyword(
        self, question: str, mismatch_item: str, content: str
    ) -> dict:
        """
        4단계(보조 경로, 2026-06-28 신규) — "답변과 다름" 항목 하나에 대해,
        _resolve_single_mismatch처럼 원본에서 정확한 발췌를 찾는 방식
        (_find_snippet_in_content 1~3단계)이 모두 실패했을 때만 호출되는
        마지막 수단.

        설계 의도:
        - 발췌-치환 방식(원문과 정확히 일치하는 부분만 좁게 바꿔치기)은
          창작이 거의 없어 안전하지만, 외부 AI 교차검증 라운드를 여러 번
          거치면서 검증 항목의 표현이 원본 문구와 점점 멀어지면, 모델이
          만든 발췌가 원문 어디에도 못 찾아지는 경우가 늘어남. 이 경우
          기존 구조는 그냥 "자동 수정 실패"로 포기했음.
        - 이 메서드는 "정확한 한 문장 발췌"가 아니라, "이 항목이 본문의
          어느 단락(여러 문장으로 이뤄진 한 덩어리)을 다루고 있는지"를
          먼저 식별하게 하고, 그 단락 전체를 새로 작성하게 함. 발췌 단위가
          넓어지는 대신, "정확히 한 글자도 안 틀려야 한다"는 제약이
          없어져서 식별 성공률이 올라감.
        - 여전히 "문서 전체를 모델이 다시 써내는" 방식(과거 B/C/D안에서
          문제가 됐던 구조)은 아님 — 식별된 단락 하나만 교체 대상이고,
          나머지 본문은 코드가 건드리지 않음. 단락 단위로 범위를 좁혀
          "일부 내용을 놓치는" 위험을 줄임.
        - 이 경로로 처리된 항목은 _apply_corrections에서 항상 "키워드
          기반 재작성 — 확인 권장"이라는 표시를 붙여, 발췌-치환 방식보다
          신뢰도가 낮을 수 있음을 투명하게 알림.

        Parameters
        ----------
        question : str
            원래 질문 (참고용)
        mismatch_item : str
            검증 결과에서 추출된 "답변과 다름" 항목 하나
        content : str
            지금까지 누적 수정된 본문 전체 (이 안에서 단락을 찾아야 함)

        Returns
        -------
        dict
            {
                "target_paragraph": str,  # 본문에서 찾은, 교체 대상이 될 단락
                                          # 전체(원문 그대로). 못 찾으면 빈 문자열.
                "new_paragraph": str,     # 그 단락을 대체할 새로 작성된 단락.
                                          # 못 만들면 빈 문자열.
                "short_reason": str,      # 1줄 요약
            }
            실패 시 모든 값이 빈 문자열인 dict.
        """
        resolve_prompt = f"""다음은 세무 자문 문서의 검증 결과 중 "답변과 다름"으로 판정된 항목
하나와, 그 문서의 본문 전체입니다. 이 항목이 본문의 어느 단락을 다루고 있는지
찾아서, 그 단락 전체를 검증 결과의 정답을 반영해 새로 작성해주세요.

[원래 질문 — 참고용]
{question}

[검증 결과 항목 — 이 안에 "원본이 뭐라고 했는지"와 "정답이 무엇인지"가
모두 들어있습니다]
{mismatch_item}

[본문 전체]
{content}

[작업]
1. 위 항목이 다루는 내용이 본문의 어느 단락에 있는지 찾으세요. 그 단락을
   "target_paragraph"에 본문에 쓰여 있는 그대로(한 글자도 바꾸지 않고)
   옮겨 적으세요. 단락은 보통 한 줄(불릿 항목 하나) 또는 연속된 몇 개
   문장 단위입니다. 너무 넓게 잡아서 관련 없는 다른 내용까지 포함하지
   마세요.
2. 검증 결과에 적힌 정답 내용을 반영해서, 그 단락을 대체할 새 단락을
   "new_paragraph"로 작성하세요. 원본과 같은 어조(전문가 답변체)와
   형식(불릿이었다면 불릿, 마크다운 볼드 표시가 있었다면 유지)을
   유지하면서, 내용만 정답에 맞게 자연스럽게 새로 쓰세요.
   - 같은 내용을 중복해서 두 번 적지 마세요.
   - 문장이 중간에 어색하게 끊기거나 두 표현이 이상하게 이어붙지 않도록,
     전체를 자연스러운 하나의 글로 다시 쓰세요.
3. 무엇을 어떻게 고쳤는지 5~15자 내의 짧은 한국어 구절로 "short_reason"에
   요약하세요.
4. 본문에서 해당 단락을 찾을 수 없거나, 정답이 불명확해 새로 쓸 수 없다면
   모든 값을 빈 문자열로 두세요. 억지로 만들어내지 마세요.

[출력 형식 — 다른 설명 없이 이 JSON 형식만 출력하세요. 마크다운 코드블록
표시(```)도 쓰지 마세요]
{{"target_paragraph": "...", "new_paragraph": "...", "short_reason": "..."}}
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=resolve_prompt,
                config=genai_types.GenerateContentConfig(temperature=0.2),
            )
            raw_text = response.text.strip() if response.candidates else ""
        except Exception as e:
            print(f"[경고] 키워드 기반 단락 재작성 호출 실패: {e}")
            return {"target_paragraph": "", "new_paragraph": "", "short_reason": ""}

        if not raw_text:
            return {"target_paragraph": "", "new_paragraph": "", "short_reason": ""}

        cleaned = raw_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("```")[1]
            if cleaned.startswith("json"):
                cleaned = cleaned[4:]
            cleaned = cleaned.strip()

        try:
            import json
            parsed = json.loads(cleaned)
        except (json.JSONDecodeError, ValueError) as e:
            print(f"[경고] 키워드 기반 단락 재작성 결과 JSON 파싱 실패: {e} / 원본: {raw_text[:200]}")
            return {"target_paragraph": "", "new_paragraph": "", "short_reason": ""}

        return {
            "target_paragraph": str(parsed.get("target_paragraph", "")),
            "new_paragraph": str(parsed.get("new_paragraph", "")),
            "short_reason": str(parsed.get("short_reason", "")),
        }

    def _recommend_knowledge_file(self, question: str, corrected_content: str) -> tuple:
        """
        지식베이스 저장 시 어느 파일이 적합한지만 가볍게 판단하는 보조 호출.
        검증/수정과는 무관한 별개의 단순 분류 작업이라 분리함.

        Returns
        -------
        tuple
            (recommended_file: str, recommended_reason: str)
        """
        options_text = ", ".join(self.KNOWLEDGE_FILE_OPTIONS)
        recommend_prompt = f"""다음 세무 질의응답을 아래 파일들 중 어디에 저장하는 게 가장 적합한지
판단하세요. 선택 가능한 파일:
{options_text}
(01_공통_세무질의회신집.txt: 특정 업종에 국한되지 않는 일반 세무 질의/회신
 02_업종별_기장유의사항.txt: 약국 등 특정 업종에 특화된 기장 유의사항
 04_부가세_처리지침.txt: 부가가치세 관련 처리 기준
 05_기타_세법_예규.txt: 법인세/소득세 등 그 외 세법 예규 및 개정사항)

[질문]
{question}

[답변 내용]
{corrected_content}

첫 줄에 파일명을 정확히 그대로 적고, 둘째 줄에 이유를 한 줄로 적으세요.
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=recommend_prompt,
                config=genai_types.GenerateContentConfig(temperature=0.2),
            )
            text = response.text if response.candidates else ""
        except Exception:
            text = ""

        recommended_file = self.KNOWLEDGE_FILE_OPTIONS[0]
        recommended_reason = "추천 사유를 확인하지 못했습니다 — 기본값으로 지정되었습니다."
        if text:
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            if lines:
                candidate = lines[0]
                for opt in self.KNOWLEDGE_FILE_OPTIONS:
                    if opt in candidate or opt.replace(".txt", "") in candidate:
                        recommended_file = opt
                        break
                if len(lines) > 1:
                    recommended_reason = lines[1]

        return recommended_file, recommended_reason

    # ------------------------------------------------------------------
    # 지식베이스 확정 저장 - v1.6 변경 (구글시트 기반으로 전환)
    # ------------------------------------------------------------------
    def confirm_to_knowledge_base(
        self,
        question: str,
        confirmed_content: str,
        target_file: str = "01_공통_세무질의회신집.txt",
    ) -> str:
        """
        회계사가 검증/확정한 질의응답을 지식베이스에 추가 저장.

        설계 의도 (2026-06-25 변경):
        - 기존에는 로컬 _knowledge 폴더의 .txt 파일에 직접 텍스트를 추가했음. 하지만
          웹 서버(Streamlit Cloud)는 GitHub 저장소의 임시 복제본이라, 서버에서
          파일에 추가한 내용은 서버 재시작/재배포 시 사라짐(PC와 동기화되지 않음).
        - 변경: 구글시트가 연동되어 있으면(SheetLogger.enabled) 그쪽의 '지식베이스'
          탭에 저장함 — PC/웹 어디서 저장하든 같은 시트에 쌓이므로 항상 동기화됨.
        - 구글시트 연동이 비활성 상태(.env/Secrets 미설정)인 경우에는 기존 방식
          그대로 로컬 _knowledge 폴더 파일에 저장함 (하위 호환, PC 단독 사용 시에도
          여전히 동작하도록).

        주의: 이 메서드를 호출하기 전에 반드시 verify_pin()으로 PIN을 검증해야 합니다.
        이 메서드 자체는 PIN 검증을 하지 않습니다 (호출 측 책임).

        Parameters
        ----------
        question : str
            원래 질문
        confirmed_content : str
            확정할 내용 (회계사가 검토/수정한 최종 텍스트)
        target_file : str
            분류 역할의 파일명 (구글시트 저장 시 '분류' 컬럼 값으로 사용,
            로컬 폴백 저장 시에는 실제 파일명으로 사용)

        Returns
        -------
        str
            저장 결과를 설명하는 문자열 (구글시트 저장 시 탭 이름, 로컬 저장 시 파일 경로)
        """
        if self.sheet_logger and self.sheet_logger.enabled:
            saved = self.sheet_logger.add_knowledge_entry(
                category=target_file.replace(".txt", ""),
                question=question,
                confirmed_content=confirmed_content,
            )
            self.knowledge_cache = None  # 캐시 무효화 - 다음 질의부터 즉시 반영되도록
            if saved:
                return f"구글시트 '{SheetLogger.KNOWLEDGE_SHEET_NAME}' 탭 (분류: {target_file.replace('.txt', '')})"
            # 구글시트 저장이 실패하면 로컬 파일로 안전하게 폴백
            print("[경고] 구글시트 지식베이스 저장 실패 — 로컬 파일로 대신 저장합니다.")

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

        print("\n조회 중입니다...\n")
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
