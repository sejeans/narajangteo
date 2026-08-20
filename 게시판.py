# -*- coding: utf-8 -*-
"""기관 자체 홈페이지 게시판 어댑터

나라장터에 안 올라오는 공고가 있다. 크로스체크 실측에서 자체 사이트 106건 중
15건(14.2%)이 나라장터에 없었다. 그 15건을 얻으려고 기관 홈페이지 목록을
직접 읽는다.

이 파일은 **수집기를 모른다.** 하는 일은 두 가지뿐이다.

  1. 목록 페이지를 긁어 나라장터 API 와 같은 모양의 dict 로 바꾼다 (fetch)
  2. 그 결과가 나라장터 공고와 같은 건인지 짝지어 준다 (짝짓기)

키 이름을 나라장터 API 그대로(bidNtceNo, bidNtceNm ...) 쓰는 이유는
수집기의 judge·snapshot·diff_notice 를 한 줄도 안 고치고 쓰기 위해서다.
없는 값은 빈 문자열로 둔다. money() 와 deadline() 은 이미 빈 값을 견딘다.

게시판마다 다른 것(주소·칸 위치·기관명·켜고 끄기)은 전부 게시판.yaml 에 있다.
파이썬을 열지 않고 게시판을 늘릴 수 있어야 한다.

  python 게시판.py                 켜져 있는 게시판을 다 읽어 목록만 보여준다
  python 게시판.py 중소기업중앙회    한 곳만 (꺼둔 곳도 강제로 읽는다)
"""

from __future__ import annotations

import re
import ssl
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urljoin

try:
    import requests
    import yaml
except ImportError:  # pragma: no cover - 수집기가 먼저 같은 말을 한다
    sys.exit("필요한 패키지를 설치하세요:\n\n    pip install requests pyyaml\n")


# 상대는 우리 고객사 홈페이지다. 차단당하면 크롤링만 끊기는 게 아니다.
# 정체를 숨기지 않는다.
# HTTP 머리글은 latin-1 로만 보낼 수 있다. 한글을 한 글자라도 넣으면
# 요청이 나가기도 전에 UnicodeEncodeError 로 죽는다. 영문으로만 적는다.
UA = ("NICEPNI-BidWatch/1.0 (bid notice watcher; contact the sender "
      "of this request to be excluded)")
TIMEOUT = 30
MAX_RETRY = 2            # 죽은 사이트를 두들기지 않는다
SLEEP = 1.2              # 요청 사이 최소 간격(초). 동시 요청은 하지 않는다

BOARDS_YAML = "게시판.yaml"

# 서버 인증서를 어떻게 검사할까. load_boards 가 게시판.yaml 의 '설정' 절을
# 읽어 여기에 넣는다. 게시판별이 아니라 이 PC 의 망 사정이라 한 군데에 둔다.
#   "자동"      윈도우에 설치된 인증서로 검사한다 (기본)
#   "경로.pem"  그 파일에 든 인증서로만 검사한다
#   False       검사하지 않는다. 마지막 수단이다
인증서 = "자동"
_세션: "requests.Session | None" = None


class _인증서어댑터(requests.adapters.HTTPAdapter):
    """인증서를 어디서 가져와 검사할지 정한다.

    두 가지를 requests 기본값과 다르게 한다.

    1. **인증서 목록을 윈도우에서 가져온다.** requests 는 certifi(파이썬이
       들고 다니는 공개 CA 목록)만 본다. 사내 망에 SSL 을 가로채는 장비가
       있으면 그 장비의 루트 인증서는 윈도우에는 깔려 있어도 certifi 에는
       없어서, 멀쩡한 사이트가 전부 인증서 오류로 막힌다.
    2. **VERIFY_X509_STRICT 를 끈다.** urllib3 2.x 가 기본으로 켜는 서식
       검사다. 가로채기 장비가 만들어 내는 인증서는 Authority Key Identifier
       같은 항목이 빠져 있는 일이 흔해 여기서 걸린다. 신뢰의 문제가 아니라
       서식의 문제라 이 검사만 끄고 나머지 검증(서명·유효기간·호스트이름)은
       그대로 둔다.
    """

    def __init__(self, cafile: str | None = None):
        self._cafile = cafile
        super().__init__()

    def init_poolmanager(self, *a, **kw):
        ctx = ssl.create_default_context(cafile=self._cafile)
        ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
        kw["ssl_context"] = ctx
        return super().init_poolmanager(*a, **kw)


def 세션() -> "requests.Session":
    """연결을 재살려 쓴다. 인증서 설정이 바뀌면 _설정적용 이 지운다."""
    global _세션
    if _세션 is None:
        _세션 = requests.Session()
        _세션.mount("https://",
                   _인증서어댑터(인증서 if isinstance(인증서, str)
                             and 인증서 != "자동" else None))
    return _세션


# ---------------------------------------------------------------------------
# 목록 HTML 읽기
#
# bs4·lxml 을 쓰지 않는다. 지금 배포는 PyInstaller 로 묶은 exe 하나뿐이라
# 의존성 하나가 곧 담당자 PC 의 재설치다. 게시판 목록은 어차피 <tr><td> 표라
# 표준 라이브러리 HTMLParser 로 충분하다.
# ---------------------------------------------------------------------------

# 화면에는 보이지만 제목이 아닌 것들. NEW 딱지, 숨김 텍스트, 아이콘 라벨.
_잡음_클래스 = re.compile(r"\b(new|blind|hidden|skip|ico|icon|tag|label)\b", re.I)


class 표읽기(HTMLParser):
    """줄마다 (칸 글자 목록, 그 줄에 있던 링크 문자열) 을 모은다.

    어느 <table> 인지 고르지 않는다. 게시판 페이지에는 표가 여러 개 있고
    (레이아웃용 표, 검색 표) 어느 것이 목록인지는 사이트마다 다르다.
    대신 '번호찾기 정규식에 걸리는 링크를 가진 줄' 만 나중에 남긴다.
    그러면 표를 고를 필요가 없어진다.

    줄과 칸이 무엇인지는 게시판마다 다르다. 표로 그린 곳이 많지만
    <li> 안에 <div> 를 늘어놓는 곳도 있다 (새마을금고중앙회). 그래서
    행태그·칸태그를 밖에서 받는다. 기본값은 표다.
    """

    def __init__(self, 행태그: str = "tr", 칸태그: str = "td,th") -> None:
        super().__init__(convert_charrefs=True)
        self._행 = {t.strip().lower() for t in 행태그.split(",") if t.strip()}
        self._칸태그 = {t.strip().lower() for t in 칸태그.split(",") if t.strip()}
        self.줄: list[tuple[list[str], str]] = []
        self._칸: list[str] = []
        self._글자: list[str] = []
        self._링크: list[str] = []
        self._칸안 = False
        self._건너뛰기 = 0

    # -- 태그 --------------------------------------------------------------
    def handle_starttag(self, tag: str, attrs: list) -> None:
        a = dict(attrs)
        if self._건너뛰기:
            self._건너뛰기 += 1
            return
        if _잡음_클래스.search(a.get("class") or ""):
            self._건너뛰기 = 1
            return
        if tag in self._행:
            self._끝칸()
            self._칸, self._링크 = [], []
        elif tag in self._칸태그:
            self._끝칸()
            self._칸안 = True
        # 링크는 <a href> 뿐 아니라 <span onclick="goView(...)"> 로도 온다.
        for k in ("href", "onclick", "data-url"):
            if a.get(k):
                self._링크.append(a[k])

    def handle_endtag(self, tag: str) -> None:
        if self._건너뛰기:
            self._건너뛰기 -= 1
            return
        if tag in self._칸태그:
            self._끝칸()
        elif tag in self._행:
            self._끝칸()
            if self._칸:
                self.줄.append((self._칸, " ".join(self._링크)))
            self._칸, self._링크 = [], []

    def handle_data(self, data: str) -> None:
        if not self._건너뛰기 and self._칸안:
            self._글자.append(data)

    # -- 안쪽 --------------------------------------------------------------
    def _끝칸(self) -> None:
        if self._칸안:
            self._칸.append(" ".join("".join(self._글자).split()))
            self._글자 = []
            self._칸안 = False


def 표줄(html: str, 행태그: str = "tr",
       칸태그: str = "td,th") -> list[tuple[list[str], str]]:
    p = 표읽기(행태그, 칸태그)
    try:
        p.feed(html)
        p.close()
    except Exception:       # noqa: BLE001 - 깨진 HTML 이라도 읽은 데까지 쓴다
        pass
    return p.줄


# ---------------------------------------------------------------------------
# 값 다듬기
# ---------------------------------------------------------------------------

_날짜 = re.compile(r"(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})")


def 날짜(글자: str) -> str:
    """'2026.08.19' '2026-8-19' '2026년 8월 19일' → '2026-08-19'. 못 읽으면 ''."""
    m = _날짜.search(글자 or "")
    if not m:
        return ""
    y, mo, d = (int(x) for x in m.groups())
    try:
        return f"{date(y, mo, d):%Y-%m-%d}"
    except ValueError:
        return ""


# 제목 칸 끝에 딸려 오는 딱지. 아이콘에 붙은 대체텍스트(alt)나 화면에는
# 안 보이게 숨겨둔 글자라 제목이 아니다. 그대로 두면 점수 계산과 짝짓기에
# 다 섞인다. 아무 낱말이나 떼면 안 되므로 실제로 본 것만 적는다.
_제목꼬리 = re.compile(r"[\s\-·]*(첨부파일|첨부|파일있음|파일|새글|NEW|New|new|"
                    r"HOT|Hot|N|중요)$")


def 제목정리(글자: str) -> str:
    앞 = None
    글자 = " ".join((글자 or "").split())
    while 앞 != 글자:
        앞 = 글자
        글자 = _제목꼬리.sub("", 글자).strip()
    return 글자


def 칸(줄: list[str], 자리: int) -> str:
    """음수 자리(-1 = 맨 뒤)를 받는다. 날짜 칸은 대개 맨 뒤라 -1 이 편하다."""
    try:
        return 줄[자리]
    except IndexError:
        return ""


# ---------------------------------------------------------------------------
# 게시판 하나
# ---------------------------------------------------------------------------


@dataclass
class Board:
    이름: str
    기관: str
    목록: str
    상세: str = ""
    번호찾기: str = r"goView\((\d+)"
    제목칸: int = 1
    날짜칸: int = -1
    # 마감일시가 목록에 나오는 게시판이 드물게 있다(우체국금융개발원).
    # 있으면 적어둔다. 없으면 그 칸은 빈칸으로 나간다.
    마감칸: int | None = None
    # 목록 한 줄과 그 안의 칸을 무엇으로 볼지. 표가 아니면 li,div 처럼 준다.
    행태그: str = "tr"
    칸태그: str = "td,th"
    제목제외: list[str] = field(default_factory=list)
    제목포함: list[str] = field(default_factory=list)
    사용: bool = True

    @classmethod
    def from_yaml(cls, d: dict) -> "Board":
        알려진 = set(cls.__dataclass_fields__)
        b = cls(**{k: v for k, v in d.items() if k in 알려진})
        b.모르는항목 = [k for k in d if k not in 알려진]
        return b


@dataclass
class 결과:
    이름: str
    공고: list[dict] = field(default_factory=list)   # 기간 안에 든 것만
    전체: int = 0                                    # 목록에서 읽어낸 줄 수
    걸러냄: int = 0                                  # 제목 필터로 뺀 수
    오류: str = ""


def _받기(url: str, 알림) -> tuple[str, str]:
    """(본문, 오류). 실패해도 예외를 올리지 않는다."""
    머리 = {"User-Agent": UA, "Accept-Language": "ko"}
    for 회차 in range(1, MAX_RETRY + 1):
        try:
            res = 세션().get(url, headers=머리, timeout=TIMEOUT,
                            verify=인증서 is not False)
            res.raise_for_status()
            # 옛 게시판은 아직 EUC-KR 이다. requests 가 ISO-8859-1 로 잘못
            # 찍어 두는 일이 있어 그때만 본문에서 다시 알아낸다.
            if not res.encoding or res.encoding.lower() == "iso-8859-1":
                res.encoding = res.apparent_encoding or "utf-8"
            return res.text, ""
        except requests.exceptions.SSLError as exc:
            # 사내 망에서 SSL 을 가로채는 장비를 쓰면 여기서 걸린다.
            # 사이트가 죽은 것이 아니므로 다시 시도해봐야 소용없다.
            return "", ("서버 인증서를 확인하지 못했습니다. 사내 망이 SSL 을 "
                        "가로채는 장비를 쓰는데 그 루트 인증서가 이 PC 에 "
                        "안 깔린 경우가 대부분입니다. 게시판.yaml 의 "
                        f"설정.인증서 를 보세요 — {str(exc)[:90]}")
        except requests.RequestException as exc:
            if 회차 == MAX_RETRY:
                return "", f"{type(exc).__name__}: {str(exc)[:80]}"
            알림(f"    [재시도 {회차}/{MAX_RETRY}] {url[:60]}")
            time.sleep(2 * 회차)
        except Exception as exc:      # noqa: BLE001
            # 게시판 한 곳 때문에 회차 전체가 죽으면 안 된다. 무엇이든
            # 오류 문자열로 바꿔 돌려주고, 수집기가 '일부실패' 로 처리한다.
            return "", f"{type(exc).__name__}: {str(exc)[:80]}"
    return "", "알 수 없음"


def fetch(board: Board, begin: date | None = None, end: date | None = None,
          알림=print) -> 결과:
    """게시판 한 곳의 목록 첫 장을 읽는다.

    begin·end 는 등록일로 자른다. 나라장터와 같은 창을 봐야 한다.
    게시판 한 장에는 몇 달치가 한꺼번에 보이므로 자르지 않으면 처음 켠 날
    수십 건이 '신규' 로 쏟아진다. 그건 신규가 아니라 우리가 안 보던 것이다.
    """
    본문, 오류 = _받기(board.목록, 알림)
    if 오류:
        return 결과(board.이름, 오류=오류)

    번호 = re.compile(board.번호찾기)
    제외 = [re.compile(p) for p in board.제목제외]
    포함 = [re.compile(p) for p in board.제목포함]

    공고, 전체, 걸러냄 = [], 0, 0
    본_번호: set[str] = set()
    for 칸들, 링크 in 표줄(본문, board.행태그, board.칸태그):
        m = 번호.search(링크)
        if not m:
            continue
        no = m.group(1)
        if no in 본_번호:        # 한 줄에 같은 링크가 두 번 걸린 게시판이 있다
            continue
        본_번호.add(no)
        제목 = 제목정리(칸(칸들, board.제목칸))
        if not 제목:
            continue
        전체 += 1
        if 포함 and not any(p.search(제목) for p in 포함):
            걸러냄 += 1
            continue
        if any(p.search(제목) for p in 제외):
            걸러냄 += 1
            continue
        # 번호찾기에 이름 붙인 괄호 (?P<고정>...) 를 넣으면 그 이름 그대로
        # 상세 주소에 {고정} 으로 끼워 쓸 수 있다. 공지로 고정된 글만 주소가
        # 다른 게시판이 있어서 필요하다.
        추가 = {k: (v or "") for k, v in m.groupdict().items()}
        등록일 = 날짜(칸(칸들, board.날짜칸))
        마감 = (날짜(칸(칸들, board.마감칸))
              if board.마감칸 is not None else "")
        if 등록일 and begin and 등록일 < f"{begin:%Y-%m-%d}":
            continue
        if 등록일 and end and 등록일 > f"{end:%Y-%m-%d}":
            continue
        공고.append(레코드(board, no, 제목, 등록일, 추가, 마감))

    return 결과(board.이름, 공고, 전체, 걸러냄)


def 레코드(board: Board, no: str, 제목: str, 등록일: str,
        추가: dict | None = None, 마감: str = "") -> dict:
    """나라장터 API 와 같은 모양. 없는 값은 빈 문자열이다.

    빈 채로 두는 것: 마감일시·배정예산·추정가격·면허제한·첨부.
    게시판 목록에 없는 값을 지어내지 않는다. 메일에서 빈칸으로 보이는 편이
    틀린 값이 채워져 있는 것보다 낫다. 출처 칸이 그 빈칸을 설명한다.
    """
    return {
        "bidNtceNo": no,
        "bidNtceNm": 제목,
        "bidNtceDt": 등록일,
        "dminsttNm": board.기관,
        "ntceInsttNm": board.기관,
        "bidNtceDtlUrl": 상세주소(board, no, 추가 or {}),
        # 목록에 날짜만 있고 시각은 없다. 시각을 지어내지 않는다.
        "bidClseDt": 마감,
        "opengDt": "",
        "asignBdgtAmt": "",
        "presmptPrce": "",
        # 게시판은 차수를 안 알려준다. 제목의 (재공고) 로만 짐작한다.
        "ntceKindNm": "재공고" if "재공고" in 제목 else "등록공고",
        "bidNtceOrd": "0",
        "게시판": board.이름,
    }


def 상세주소(board: Board, no: str, 추가: dict) -> str:
    """글 하나의 주소. 못 만들면 목록 주소라도 준다.

    링크가 없는 행보다 목록으로 가는 링크가 낫다. 담당자가 거기서 제목으로
    찾을 수 있다.
    """
    if not board.상세:
        return board.목록
    try:
        return urljoin(board.목록, board.상세.format(**{"번호": no, **추가}))
    except (KeyError, IndexError):
        return board.목록


def load_boards(path: Path, 알림=print) -> list[Board]:
    if not path.exists():
        return []
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:
        알림(f"[주의] {path.name} 을 읽지 못했습니다: {exc}")
        return []
    _설정적용(raw.get("설정") or {}, path, 알림)
    목록 = []
    for d in (raw.get("게시판") or []):
        if not isinstance(d, dict) or not d.get("이름"):
            continue
        try:
            b = Board.from_yaml(d)
        except TypeError as exc:
            알림(f"[주의] 게시판 '{d.get('이름')}' 설정이 잘못됐습니다: {exc}")
            continue
        if b.모르는항목:
            알림(f"[주의] 게시판 '{b.이름}' 에 모르는 항목이 있어 무시합니다: "
               f"{', '.join(b.모르는항목)}")
        목록.append(b)
    return 목록


def _설정적용(설정: dict, path: Path, 알림) -> None:
    """게시판.yaml 의 '설정' 절. 지금은 인증서 이야기뿐이다."""
    global 인증서, _세션
    _세션 = None                  # 설정이 바뀌었으니 연결을 다시 만든다
    값 = 설정.get("인증서", "자동")
    if 값 is False:
        알림("[주의] 게시판 조회에서 서버 인증서를 검사하지 않습니다"
           " (게시판.yaml 의 설정.인증서: false).")
        인증서 = False
        return
    if isinstance(값, str) and 값.strip() and 값.strip() != "자동":
        pem = Path(값.strip())
        if not pem.is_absolute():
            pem = path.parent / pem
        if pem.exists():
            인증서 = str(pem)
            return
        알림(f"[주의] 설정.인증서 파일이 없어 '자동' 으로 돌립니다: {pem}")
    인증서 = "자동"


# ---------------------------------------------------------------------------
# 같은 공고인가 — 나라장터 건과 짝짓기
#
# 실측에서 자체사이트 106건 중 91건(86%)이 나라장터에도 있었다. 겹치는 것이
# 정상이고 안 겹치는 것이 예외다. 공고번호가 서로 달라 번호로는 못 맞춘다.
# ---------------------------------------------------------------------------

같음_컷 = 0.85
애매_컷 = 0.70
날짜폭 = 14              # 등록일이 이보다 벌어지면 다른 건으로 본다

# 두 채널이 서로 다르게 붙이는 머리말·꼬리말. 뜻을 바꾸지 않는 것만 뗀다.
# '선정' '공모' 처럼 뜻을 지고 있는 낱말은 그대로 둔다. 지우면 서로 다른
# 공고끼리도 비슷해져 엉뚱한 짝이 생긴다.
_군더더기 = re.compile(
    r"입찰\s*재?공고|재\s*공고|변경\s*공고|취소\s*공고|사업\s*공고|"
    r"입찰\s*공고|제안\s*요청서?|공고|긴급|"
    r"연구\s*용역|위탁\s*용역|일반\s*용역|용역|"
    r"\d{4}\s*년도?|제?\s*\d+\s*차")
_남길것 = re.compile(r"[0-9A-Za-z가-힣]+")


def 정규화(제목: str, 기관: str = "") -> str:
    """짝짓기용 제목. 사람이 읽을 값이 아니다.

    나라장터 '노란우산 광고제작 및 매체대행 용역'
    게시판   '[입찰공고] 노란우산 광고제작 및 매체대행'
    둘 다 '노란우산광고제작및매체대행' 이 된다.

    기관 이름은 떼어낸다. 한쪽만 제목에 기관 이름을 적는 일이 아주 흔한데
    그대로 두면 같은 공고가 안 닮은 것으로 나온다. 실제로 이것 때문에
    '과학기술인공제회 위탁운용사 선정 정량평가 대행 용역'(나라장터)과
    '위탁운용사 선정 정량평가 대행 용역'(게시판)이 갈렸다.
    다 지워 빈 문자열이 되는 경우(제목이 기관 이름뿐)는 그냥 둔다.
    """
    키 = "".join(_남길것.findall(_군더더기.sub(" ", 제목 or "")))
    if 기관:
        기관키 = "".join(_남길것.findall(_군더더기.sub(" ", 기관)))
        if 기관키 and 기관키 in 키 and len(키) > len(기관키):
            키 = 키.replace(기관키, "")
    return 키


def 닮음(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def _기관같음(기관: str, raw: dict) -> bool:
    """수요기관이든 공고기관이든 한쪽에 이름이 들어 있으면 같은 곳으로 본다.

    나라장터는 '중소기업중앙회 총무부' 처럼 부서까지 붙고, 대행사가 대신
    올리면 공고기관이 아예 다른 회사다. 그래서 한쪽만 맞아도 통과시킨다.
    """
    if not 기관:
        return True
    이름 = f"{raw.get('dminsttNm') or ''} {raw.get('ntceInsttNm') or ''}"
    return 기관 in 이름 or any(x and x in 기관 for x in 이름.split())


def _날짜가까움(a: str, b: str) -> bool:
    if not a or not b:
        return True          # 한쪽을 모르면 날짜로는 안 거른다
    try:
        d1 = datetime.strptime(a[:10], "%Y-%m-%d")
        d2 = datetime.strptime(b[:10], "%Y-%m-%d")
    except ValueError:
        return True
    return abs(d1 - d2) <= timedelta(days=날짜폭)


def 짝짓기(사이트: list[dict], 나라: list[dict],
        기관: str = "") -> list[tuple[str, dict | None, float]]:
    """사이트 공고마다 ('양쪽'|'애매'|'사이트', 짝지어진 나라장터 공고, 닮음).

    1:1 로 짝짓는다. 한 나라장터 건에 사이트 재공고 여러 건이 몰려 붙으면
    실제로는 남아 있어야 할 건이 통째로 사라진다. 실측 때 이걸 안 해서
    한쪽에만 있는 것처럼 나온 적이 있다.

    애매한 구간(0.70~0.85)은 **버리지 않는다.** 중복은 담당자가 보고 넘기면
    그만이지만 잘못 버리면 영영 모른다. 수집기가 취소·변경을 다시 알리는
    것과 같은 원칙이다.
    """
    후보 = [r for r in 나라 if _기관같음(기관, r)]
    s키 = [정규화(r.get("bidNtceNm") or "", 기관) for r in 사이트]
    n키 = [정규화(r.get("bidNtceNm") or "", 기관) for r in 후보]

    쌍 = []
    for i, r in enumerate(사이트):
        for j, q in enumerate(후보):
            if not _날짜가까움(str(r.get("bidNtceDt") or ""),
                          str(q.get("bidNtceDt") or "")):
                continue
            점 = 닮음(s키[i], n키[j])
            if 점 >= 애매_컷:
                쌍.append((점, i, j))

    쌍.sort(key=lambda x: (-x[0], x[1], x[2]))
    맞춤: dict[int, tuple[int, float]] = {}
    쓴것: set[int] = set()
    for 점, i, j in 쌍:
        if i in 맞춤 or j in 쓴것:
            continue
        맞춤[i] = (j, 점)
        쓴것.add(j)

    판정 = []
    for i in range(len(사이트)):
        if i not in 맞춤:
            판정.append(("사이트", None, 0.0))
            continue
        j, 점 = 맞춤[i]
        판정.append((("양쪽" if 점 >= 같음_컷 else "애매"), 후보[j], 점))
    return 판정


# ---------------------------------------------------------------------------
# 혼자 돌려보기 — 수집도 메일도 하지 않고 게시판만 읽는다
# ---------------------------------------------------------------------------


def _main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:      # pragma: no cover
        pass
    here = Path(__file__).resolve().parent
    boards = load_boards(here / BOARDS_YAML)
    if not boards:
        print(f"{BOARDS_YAML} 에 게시판이 없습니다.")
        return 1
    골라 = argv[1] if len(argv) > 1 else ""
    돌린것 = 0
    for b in boards:
        if 골라 and 골라 not in b.이름:
            continue
        if not 골라 and not b.사용:
            print(f"[꺼둠] {b.이름}")
            continue
        if 돌린것:
            time.sleep(SLEEP)
        돌린것 += 1
        r = fetch(b)
        if r.오류:
            print(f"[실패] {b.이름} — {r.오류}")
            continue
        print(f"\n[{b.이름}] 목록 {r.전체}건 · 제목필터로 뺌 {r.걸러냄}건 "
              f"· 남음 {len(r.공고)}건")
        for x in r.공고:
            print(f"  {x['bidNtceDt'] or '날짜없음':10}  {x['bidNtceNm'][:60]}")
            print(f"              {x['bidNtceDtlUrl']}")
    if not 돌린것:
        print("돌린 게시판이 없습니다.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(_main(sys.argv))
