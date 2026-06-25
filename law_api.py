# -*- coding: utf-8 -*-
"""
law_api.py
==========
법제처(국가법령정보 공동활용) Open API 연동 모듈

설계 근거: 실제 OC 키로 호출하여 확인한 응답 구조를 기준으로 작성됨.
(2026-06-24 사전 점검 결과 반영. 이전 버전의 영문 태그(JO_INFO, JO_NUM) 가정은
 실제 응답과 맞지 않아 발생한 오류였음 — 실제 응답은 한글 태그 사용.)

호출 흐름 (2단계):
    1. lawSearch.do  : 법령명으로 검색 → 법령일련번호(MST) 획득
    2. lawService.do : MST로 본문(조문 전체) 조회

응답 구조 요약 (실측):
    <LawSearch>
        <law id="1">
            <법령일련번호>276117</법령일련번호>
            <법령명한글><![CDATA[부가가치세법]]></법령명한글>
            ...
        </law>
    </LawSearch>

    <법령>
        <기본정보>...</기본정보>
        <조문>
            <조문단위 조문키="...">
                <조문번호>1</조문번호>
                <조문여부>조문</조문여부>   (※ "전문"인 경우는 장/절 제목이므로 제외)
                <조문제목><![CDATA[목적]]></조문제목>
                <조문내용><![CDATA[제1조(목적) ...]]></조문내용>
                <항>
                    <호>
                        <호번호><![CDATA[1.]]></호번호>
                        <호내용><![CDATA[1. "재화"란 ...]]></호내용>
                    </호>
                </항>
            </조문단위>
        </조문>
    </법령>

Windows 환경 기준으로 작성됨. 표준 라이브러리만 사용 (추가 설치 불필요).
"""

import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field


SEARCH_URL = "https://www.law.go.kr/DRF/lawSearch.do"
SERVICE_URL = "https://www.law.go.kr/DRF/lawService.do"


class LawAPIError(Exception):
    """법제처 API 호출/응답 처리 중 발생하는 오류"""
    pass


@dataclass
class LawSearchResult:
    """법령 검색 결과 1건"""
    mst: str            # 법령일련번호 (본문 조회 시 사용)
    law_id: str          # 법령ID
    name: str            # 법령명한글
    law_type: str = ""   # 법령구분명 (법률/시행령/시행규칙 등)


@dataclass
class LawArticle:
    """법령 조문 1건"""
    number: str          # 조문번호
    title: str            # 조문제목
    content: str          # 조문내용
    sub_items: list = field(default_factory=list)  # 호 내용 리스트


class LawAPIClient:
    """법제처 Open API 클라이언트"""

    def __init__(self, oc_key: str, timeout: int = 10):
        if not oc_key or not oc_key.strip():
            raise ValueError("[오류] 법제처 OC 인증키가 비어 있습니다.")
        self.oc_key = oc_key.strip()
        self.timeout = timeout

    # ------------------------------------------------------------------
    def _fetch(self, url: str) -> str:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except urllib.error.URLError as e:
            raise LawAPIError(f"법제처 API 네트워크 오류: {e}")

    @staticmethod
    def _parse_xml(xml_text: str) -> ET.Element:
        try:
            return ET.fromstring(xml_text)
        except ET.ParseError as e:
            raise LawAPIError(
                f"법제처 API 응답을 XML로 해석할 수 없습니다: {e}\n"
                f"(OC 키 오류 또는 법제처 서버 문제일 수 있습니다)"
            )

    # ------------------------------------------------------------------
    def search_law(self, query: str, display: int = 5) -> list:
        """
        법령명으로 검색하여 법령일련번호(MST) 등을 반환

        Parameters
        ----------
        query : str
            검색할 법령명 (예: "부가가치세법")
        display : int
            반환받을 결과 개수 (최대 100)

        Returns
        -------
        list[LawSearchResult]
        """
        url = SEARCH_URL + "?" + urllib.parse.urlencode({
            "OC": self.oc_key,
            "target": "law",
            "type": "XML",
            "query": query,
            "display": display,
        })

        xml_text = self._fetch(url)
        root = self._parse_xml(xml_text)

        result_code = root.findtext("resultCode", "")
        if result_code and result_code != "00":
            result_msg = root.findtext("resultMsg", "알 수 없는 오류")
            raise LawAPIError(f"법제처 API 오류 응답 (코드 {result_code}): {result_msg}")

        results = []
        for law_node in root.findall("law"):
            mst = law_node.findtext("법령일련번호", "").strip()
            law_id = law_node.findtext("법령ID", "").strip()
            name = (law_node.findtext("법령명한글", "") or "").strip()
            law_type = (law_node.findtext("법령구분명", "") or "").strip()
            if mst:
                results.append(LawSearchResult(mst=mst, law_id=law_id, name=name, law_type=law_type))

        return results

    # ------------------------------------------------------------------
    def get_law_articles(self, mst: str) -> list:
        """
        법령일련번호(MST)로 본문 조문 전체를 조회

        Parameters
        ----------
        mst : str
            법령일련번호 (search_law 결과의 mst 값)

        Returns
        -------
        list[LawArticle]
            "전문"(장/절 제목 등 비조문 항목)은 제외하고 실제 조문만 반환
        """
        url = SERVICE_URL + "?" + urllib.parse.urlencode({
            "OC": self.oc_key,
            "target": "law",
            "type": "XML",
            "MST": mst,
        })

        xml_text = self._fetch(url)
        root = self._parse_xml(xml_text)

        articles = []
        for jomun in root.findall(".//조문단위"):
            jomun_type = (jomun.findtext("조문여부", "") or "").strip()
            if jomun_type != "조문":
                # "전문" 등 장/절 제목은 실제 조문이 아니므로 건너뜀
                continue

            number = (jomun.findtext("조문번호", "") or "").strip()
            title = (jomun.findtext("조문제목", "") or "").strip()
            content = (jomun.findtext("조문내용", "") or "").strip()

            sub_items = []
            for ho in jomun.findall(".//호"):
                ho_content = (ho.findtext("호내용", "") or "").strip()
                if ho_content:
                    sub_items.append(ho_content)

            articles.append(LawArticle(
                number=number,
                title=title,
                content=content,
                sub_items=sub_items,
            ))

        return articles

    # ------------------------------------------------------------------
    # 국세청 법령해석(질의회신/예규) 목록 검색 - v1.3 추가
    # ------------------------------------------------------------------
    def search_nts_interpretations(self, query: str, display: int = 3, search: int = 2) -> list:
        """
        국세청 법령해석(질의회신/예규) 목록을 검색

        중요: 이 API는 "목록"만 제공합니다. 법제처는 국세청 법령해석에 대해
        본문 조회 API를 제공하지 않습니다 (2026-06-24 실측 확인: 신청 화면의
        "국세청 법령해석" 행에는 본문 HTML/XML/JSON 칸이 존재하지 않음).
        따라서 안건명/안건번호/해석일자/링크까지만 가져올 수 있고,
        실제 회신 본문은 별도 방법(Google Search 등)으로 보강해야 합니다.

        검색 동작 관련 중요 실측 사항 (2026-06-24 확인):
        - 단일 단어 검색은 잘 작동하지만, 여러 단어를 쉼표/공백으로 합친 구문은
          0건이 나옴 (안건명에 모든 단어가 포함되어야 하는 AND/구문일치 방식으로 추정).
          → 호출 측에서 핵심 단어 1개만 넘기는 것을 권장.
        - search=1(기본, 안건명 검색)보다 search=2(본문검색)가 훨씬 많은 결과를 반환함
          (예: "음식점업" 안건명검색 29건 vs 본문검색 119건). 기본값을 2로 설정함.

        실측 응답 구조:
            <CgmExpc>
                <resultCode>00</resultCode>
                <cgmExpc id="1">
                    <법령해석일련번호>320330</법령해석일련번호>
                    <안건명><![CDATA[...]]></안건명>
                    <안건번호><![CDATA[서면-2020-법규재산-4752[법규과-3354]]]></안건번호>
                    <해석기관명>국세청</해석기관명>
                    <해석일자>2022.11.21</해석일자>
                    <법령해석상세링크>https://taxlaw.nts.go.kr/...</법령해석상세링크>
                </cgmExpc>
            </CgmExpc>

        Parameters
        ----------
        query : str
            검색어. 핵심 단어 1개를 권장 (복합 구문은 0건이 나올 위험이 큼).
        display : int
            반환받을 결과 개수 (최대 100)
        search : int
            검색 범위. 1=안건명, 2=본문(기본값, 더 많은 결과 반환)

        Returns
        -------
        list[dict]
            [{"id": str, "title": str, "case_no": str, "date": str, "link": str}, ...]
        """
        url = SEARCH_URL + "?" + urllib.parse.urlencode({
            "OC": self.oc_key,
            "target": "ntsCgmExpc",
            "type": "XML",
            "query": query,
            "display": display,
            "search": search,
        })

        xml_text = self._fetch(url)
        root = self._parse_xml(xml_text)

        result_code = root.findtext("resultCode", "")
        if result_code and result_code != "00":
            result_msg = root.findtext("resultMsg", "알 수 없는 오류")
            raise LawAPIError(f"법제처 API 오류 응답 (코드 {result_code}): {result_msg}")

        results = []
        for node in root.findall("cgmExpc"):
            interp_id = (node.findtext("법령해석일련번호", "") or "").strip()
            title = (node.findtext("안건명", "") or "").strip()
            case_no = (node.findtext("안건번호", "") or "").strip()
            date = (node.findtext("해석일자", "") or "").strip()
            link = (node.findtext("법령해석상세링크", "") or "").strip()
            if interp_id:
                results.append({
                    "id": interp_id,
                    "title": title,
                    "case_no": case_no,
                    "date": date,
                    "link": link,
                })

        return results

    # ------------------------------------------------------------------
    def get_law_text_by_name(self, law_name: str, max_articles: int = None) -> str:
        """
        법령명으로 검색 후 바로 조문 전체를 텍스트로 반환 (편의 메서드)

        Parameters
        ----------
        law_name : str
            법령명 (예: "부가가치세법")
        max_articles : int, optional
            반환할 최대 조문 수 (None이면 전체)

        Returns
        -------
        str
            "[법령명] (법령구분)\n\n제1조(제목)\n내용\n  - 호내용\n\n제2조..." 형식의 텍스트.
            검색 결과가 없으면 빈 문자열 반환.
        """
        search_results = self.search_law(law_name, display=1)
        if not search_results:
            return ""

        target = search_results[0]
        articles = self.get_law_articles(target.mst)

        if max_articles:
            articles = articles[:max_articles]

        lines = [f"[{target.name}] ({target.law_type})\n"]
        for art in articles:
            header = f"제{art.number}조" + (f"({art.title})" if art.title else "")
            lines.append(header)
            lines.append(art.content)
            for sub in art.sub_items:
                lines.append(f"  - {sub}")
            lines.append("")

        return "\n".join(lines)


# ----------------------------------------------------------------------
# 단독 실행 테스트
# ----------------------------------------------------------------------
if __name__ == "__main__":
    import sys

    print("=" * 60)
    print("law_api.py 단독 실행 테스트")
    print("=" * 60)

    oc = input("\n법제처 OC 인증키를 입력하세요: ").strip()
    if not oc:
        print("OC 키가 입력되지 않아 종료합니다.")
        sys.exit(1)

    client = LawAPIClient(oc_key=oc)

    query = input("검색할 법령명을 입력하세요 (예: 부가가치세법): ").strip() or "부가가치세법"

    print(f"\n[검색 중] '{query}' ...")
    try:
        results = client.search_law(query, display=5)
    except LawAPIError as e:
        print(str(e))
        sys.exit(1)

    if not results:
        print("검색 결과가 없습니다.")
        sys.exit(0)

    print(f"\n검색 결과 {len(results)}건:")
    for r in results:
        print(f"  - {r.name} ({r.law_type}) [MST={r.mst}]")

    first = results[0]
    print(f"\n[본문 조회 중] {first.name} (MST={first.mst}) ...")
    try:
        articles = client.get_law_articles(first.mst)
    except LawAPIError as e:
        print(str(e))
        sys.exit(1)

    print(f"\n총 {len(articles)}개 조문 확인. 처음 3개만 출력합니다.\n")
    for art in articles[:3]:
        print(f"제{art.number}조" + (f"({art.title})" if art.title else ""))
        print(art.content)
        for sub in art.sub_items:
            print(f"  - {sub}")
        print()
