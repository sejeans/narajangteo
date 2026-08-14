# -*- coding: utf-8 -*-
"""config.yaml · 점수표.yaml 를 브라우저에서 고치는 편집기.

    python 편집기.py            (또는 편집기.bat 더블클릭)

메모장으로 YAML 을 직접 고치면 들여쓰기와 따옴표를 틀리기 쉽고, 고친 결과가
점수에 어떻게 먹히는지는 돌려봐야 알 수 있다. 이 편집기는 값을 고르는 화면을
주고, 저장하기 전에 공고명을 넣어 점수가 어떻게 나오는지 바로 보여준다.

주석을 지우지 않는다.
YAML 전체를 다시 써서 덮는 방식이 아니라, 고칠 줄만 찾아 그 줄만 바꾼다.
두 파일은 절반이 주석이라 덮어쓰면 설명이 통째로 날아간다.
쓰기 직전에 결과를 다시 파싱해서 의도한 값과 같은지 확인하고,
다르면 아예 쓰지 않는다.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    import yaml
except ImportError:
    sys.exit("필요한 패키지를 설치하세요:\n\n    pip install pyyaml\n")

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

CONFIG_PATH = HERE / "config.yaml"
TABLE_PATH = HERE / "점수표.yaml"
BACKUP_DIR = HERE / "백업"
VERIFY_SCRIPT = HERE / "점수표_검증.py"

# SMTP 비밀번호는 화면에 그대로 띄우지 않는다. 사내 메일 계정 자격증명이다.
# 저장할 때 이 값이 그대로 돌아오면 '고치지 않았다' 로 본다.
#
# data.go.kr 인증키는 가리지 않는다. 요금이 붙지 않는 무료 API 이고,
# 어차피 config.yaml 에 평문으로 들어 있고 화면은 이 PC 에서만 열린다.
# 가려놓으면 지금 들어있는 키가 맞는지 확인할 방법이 없어 불편하기만 했다.
MASK = "•••••••• (그대로 둠)"


# ===========================================================================
#  YAML 줄 단위 편집 — 주석을 살리기 위한 것
# ===========================================================================

KEY_RE = re.compile(r"^( *)([^\s#\-][^:]*):(.*)$")
ITEM_RE = re.compile(r"^( *)-\s*(.*)$")

# 따옴표 없이 그대로 쓸 수 있는 문자열. 첫 글자에 YAML 특수문자가 오면 안 된다.
PLAIN_RE = re.compile(r"^[^\s#&*!|>'\"%@`\[\]{},:\-?][^#:]*$")


def is_skippable(line: str) -> bool:
    """빈 줄이거나 주석 줄."""
    s = line.strip()
    return s == "" or s.startswith("#")


def fmt(v) -> str:
    """파이썬 값을 YAML 한 조각으로."""
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return "''"
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v)
    if s == "":
        return "''"
    # '3865' 처럼 숫자로 읽힐 문자열은 반드시 따옴표를 씌운다.
    try:
        float(s)
        return "'" + s.replace("'", "''") + "'"
    except ValueError:
        pass
    if s.lower() in ("true", "false", "yes", "no", "on", "off", "null", "~"):
        return "'" + s.replace("'", "''") + "'"
    if PLAIN_RE.match(s) and not s.endswith(":") and ": " not in s:
        return s
    return "'" + s.replace("'", "''") + "'"


def parse_scalar(text: str):
    """YAML 한 조각을 파이썬 값으로. 못 읽으면 원문 그대로."""
    try:
        v = yaml.safe_load(text)
    except yaml.YAMLError:
        return text
    if v is None and text.strip() not in ("", "~", "null", "Null", "NULL"):
        return text
    return v


def split_comment(rest: str) -> tuple[str, str]:
    """키 줄의 값 부분에서 꼬리 주석을 떼어낸다.

    따옴표 안의 # 은 주석이 아니다.  " '' # 예: mail.co.kr " 같은 줄이 있다.
    """
    quote = None
    for i, ch in enumerate(rest):
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == "#" and (i == 0 or rest[i - 1] in " \t"):
            return rest[:i], rest[i:]
    return rest, ""


class YamlDoc:
    """줄 목록을 들고 있다가, 고쳐야 하는 줄만 바꿔 쓴다."""

    def __init__(self, path: Path):
        self.path = path
        # newline="" 이 없으면 파이썬이 CRLF 를 LF 로 바꿔서 읽어버린다.
        # 그러면 줄바꿈을 잘못 판단해 파일 전체가 LF 로 바뀐다.
        # config.yaml 은 LF, 점수표.yaml 은 CRLF 로 서로 다르다.
        raw = path.read_text(encoding="utf-8", newline="")
        self.nl = "\r\n" if "\r\n" in raw else "\n"
        self.trailing_nl = raw.endswith("\n")
        body = raw.replace("\r\n", "\n")
        self.lines = body.split("\n")
        if self.trailing_nl and self.lines and self.lines[-1] == "":
            self.lines.pop()
        self.data = yaml.safe_load(body) or {}
        self.changes: list[str] = []

    # ---- 읽기 --------------------------------------------------------
    def get(self, path: str):
        cur = self.data
        for p in path.split("."):
            if not isinstance(cur, dict):
                return None
            cur = cur.get(p)
        return cur

    def _put(self, path: str, value) -> None:
        parts = path.split(".")
        cur = self.data
        for p in parts[:-1]:
            cur = cur.setdefault(p, {})
        cur[parts[-1]] = value

    # ---- 위치 찾기 ---------------------------------------------------
    def _find_key(self, key: str, start: int, end: int) -> int | None:
        """[start, end) 안에서 가장 얕은 들여쓰기의 형제 키 중 key 인 줄."""
        base = None
        for i in range(start, end):
            line = self.lines[i]
            if is_skippable(line):
                continue
            m = KEY_RE.match(line)
            if not m:
                continue
            ind = len(m.group(1))
            if base is None or ind < base:
                base = ind
            if ind != base:
                continue
            if m.group(2).strip().strip("\"'") == key:
                return i
        return None

    def key_line(self, path: str) -> int:
        start, end = 0, len(self.lines)
        idx = None
        for p in path.split("."):
            idx = self._find_key(p, start, end)
            if idx is None:
                raise KeyError(f"{self.path.name} 에서 {path} 를 찾지 못했습니다")
            start, end = self.block_range(idx)
        return idx

    def block_range(self, idx: int) -> tuple[int, int]:
        """키 줄 idx 의 자식들이 차지하는 [start, end).

        중간의 빈 줄·주석은 자식 안쪽으로 보되, 마지막 자식 뒤의 주석은
        다음 항목을 설명하는 것이므로 넣지 않는다.
        """
        m = KEY_RE.match(self.lines[idx])
        key_ind = len(m.group(1)) if m else 0
        start = idx + 1
        end = start
        for i in range(start, len(self.lines)):
            line = self.lines[i]
            if is_skippable(line):
                continue
            if len(line) - len(line.lstrip(" ")) <= key_ind:
                break
            end = i + 1
        return start, end

    def items(self, idx: int) -> tuple[list[tuple[int, str]], int]:
        """블록 리스트의 (줄번호, 값 원문) 목록과 그 들여쓰기 폭."""
        start, end = self.block_range(idx)
        out: list[tuple[int, str]] = []
        base = None
        for i in range(start, end):
            line = self.lines[i]
            if is_skippable(line):
                continue
            m = ITEM_RE.match(line)
            if not m:
                continue
            ind = len(m.group(1))
            if base is None:
                base = ind
            if ind != base:
                continue
            out.append((i, m.group(2)))
        return out, (base if base is not None else 0)

    # ---- 쓰기 --------------------------------------------------------
    def set_scalar(self, path: str, value, label: str = "") -> None:
        old = self.get(path)
        if old == value:
            return
        idx = self.key_line(path)
        m = KEY_RE.match(self.lines[idx])
        ind, key, rest = m.group(1), m.group(2), m.group(3)
        val_text, comment = split_comment(rest)

        multiline = isinstance(value, str) and "\n" in value
        was_block = val_text.strip() in ("|", "|-", "|+", ">", ">-", ">+")

        if multiline or was_block:
            start, end = self.block_range(idx)
            child = ind + "  "
            body = [(child + ln if ln else "") for ln in str(value).split("\n")]
            # 블록 스칼라 헤더 뒤에는 주석을 붙이지 않는다. 헷갈리기 쉽다.
            self.lines[start:end] = body
            self.lines[idx] = f"{ind}{key}: |-"
        else:
            line = f"{ind}{key}: {fmt(value)}"
            if comment.strip():
                line += "  " + comment.strip()
            self.lines[idx] = line

        self._put(path, value)
        self.changes.append(f"{label or path}: {brief(old)} → {brief(value)}")

    def set_list(self, path: str, values: list, label: str = "") -> None:
        old = list(self.get(path) or [])
        if old == values:
            return
        idx = self.key_line(path)
        m = KEY_RE.match(self.lines[idx])
        ind, key, rest = m.group(1), m.group(2), m.group(3)
        val_text, comment = split_comment(rest)
        inline = val_text.strip().startswith("[")

        if inline or not values:
            # cc: [] 처럼 한 줄로 쓰여 있거나, 전부 지워 빈 목록이 된 경우.
            # 빈 목록을 블록으로 두면 값이 없는 키가 되어 None 이 된다.
            start, end = self.block_range(idx)
            if not inline:
                del self.lines[start:end]
            flow = "[" + ", ".join(fmt(v) for v in values) + "]"
            line = f"{ind}{key}: {flow}"
            if comment.strip():
                line += "  " + comment.strip()
            self.lines[idx] = line
        else:
            # 지운 것은 그 줄만 없애고, 새로 넣은 것은 마지막 항목 뒤에 붙인다.
            # 이렇게 하면 목록 안의 '# --- PDF 원문에서 확인됨 ---' 같은
            # 구분 주석이 그대로 남는다.
            rows, _ = self.items(idx)
            drop = [i for i, text in rows if parse_scalar(text) not in values]
            for i in sorted(drop, reverse=True):
                del self.lines[i]

            idx = self.key_line(path)
            rows, base = self.items(idx)
            have = [parse_scalar(t) for _, t in rows]
            add = [v for v in values if v not in have]
            if add:
                pos = rows[-1][0] + 1 if rows else idx + 1
                pad = " " * (base if rows else len(ind) + 2)
                self.lines[pos:pos] = [f"{pad}- {fmt(v)}" for v in add]

        self._put(path, values)
        added = [v for v in values if v not in old]
        removed = [v for v in old if v not in values]
        bits = []
        if added:
            bits.append("추가 " + ", ".join(map(str, added[:6]))
                        + (f" 외 {len(added) - 6}개" if len(added) > 6 else ""))
        if removed:
            bits.append("삭제 " + ", ".join(map(str, removed[:6]))
                        + (f" 외 {len(removed) - 6}개" if len(removed) > 6 else ""))
        self.changes.append(f"{label or path}: " + " · ".join(bits))

    def text(self) -> str:
        out = self.nl.join(self.lines)
        if self.trailing_nl:
            out += self.nl
        return out

    def save(self) -> Path:
        """다시 파싱해 의도한 값과 같은지 확인한 뒤에만 쓴다."""
        text = self.text()
        try:
            got = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise ValueError(f"고친 결과가 YAML 로 읽히지 않습니다: {exc}") from exc
        if got != self.data:
            raise ValueError("고친 결과가 의도한 값과 다릅니다: "
                             + (diff_keys(self.data, got) or "원인 미확인"))

        BACKUP_DIR.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = BACKUP_DIR / f"{self.path.stem}_{stamp}{self.path.suffix}"
        shutil.copy2(self.path, backup)
        # newline="" 이라야 위에서 정한 줄바꿈이 그대로 나간다.
        # config.yaml 은 LF, 점수표.yaml 은 CRLF 로 서로 다르다.
        with self.path.open("w", encoding="utf-8", newline="") as f:
            f.write(text)
        return backup


def brief(v, limit: int = 40) -> str:
    if isinstance(v, list):
        return f"{len(v)}개"
    s = "(비어 있음)" if v in ("", None) else str(v)
    s = s.replace("\n", " ")
    return s if len(s) <= limit else s[:limit] + "…"


def diff_keys(want, got, prefix: str = "") -> str:
    """어느 키가 어긋났는지 한 줄로."""
    if isinstance(want, dict) and isinstance(got, dict):
        for k in want:
            if k not in got:
                return f"{prefix}{k} 가 사라졌습니다"
            msg = diff_keys(want[k], got[k], f"{prefix}{k}.")
            if msg:
                return msg
        for k in got:
            if k not in want:
                return f"{prefix}{k} 가 늘었습니다"
        return ""
    if want != got:
        return f"{prefix.rstrip('.')} — 넣으려던 값 {brief(want)} / 읽힌 값 {brief(got)}"
    return ""


# ===========================================================================
#  화면에 무엇을 어떻게 보여줄지
# ===========================================================================

def F(path, label, type="str", **kw):
    d = {"path": path, "label": label, "type": type}
    d.update(kw)
    return d


TABLE_SCHEMA = [
    {"group": "문턱과 배점",
     "note": "공고명 점수가 이 문턱을 넘으면 수집·검토로 올라갑니다. "
             "문턱을 낮추면 무관한 공고가 늘고, 높이면 놓칩니다.",
     "fields": [
         F("cut_collect", "수집 문턱", "int",
           help="이 점수 이상이면 자동 수집(A·B). 기본 13"),
         F("cut_review", "검토 문턱", "int",
           help="이 점수 이상이면 검토 필요(C, 엑셀에서 회색 줄). 기본 7"),
         F("cut_review_low", "연기금·공제회 검토 문턱", "int",
           help="config.yaml 의 institutions_low_cut 에 걸리는 기관에만 적용. 기본 5"),
         F("inst_bonus", "기관 가점", "int",
           help="config.yaml 의 institutions 에 걸리면 더하는 점수. 기본 3"),
         F("inst_extra_bonus", "기관 보강 가점", "int",
           help="아래 inst_extra 목록에 걸리면 더하는 점수. 기본 2"),
     ]},
    {"group": "점수를 올리는 키워드",
     "note": "공고명에 이 낱말이 있으면 점수가 올라갑니다. "
             "괄호·공백·가운뎃점은 비교 전에 지우므로 '공정가치 평가' 와 "
             "'공정가치평가' 는 같게 봅니다.",
     "fields": [
         F("core", "core — 평가 확정어", "list", score="8점씩 (최대 3개)",
           help="이 표현이 있으면 채권평가회사 업무로 봅니다. 아래 거부권도 무효화합니다."),
         F("market", "market — 채권시장 제도·연구", "list", score="8점",
           help="'평가' 같은 행위어가 없어도 우리 영역인 주제. "
                "예탁결제원 장외 채권거래 건이 여기 걸립니다."),
         F("target", "target — 평가 대상", "list", score="2점씩 (최대 2개)",
           help="모펀드·위탁운용사는 core 가 아니라 여기 둡니다. "
                "core 에 넣으면 '운용사 선정' 공고까지 거부권이 풀립니다."),
         F("action", "action — 평가·산출 행위", "list", score="2점씩 (최대 2개)",
           help="target 과 action 이 둘 다 걸리면 +4 가 더 붙습니다."),
     ]},
    {"group": "거부권 — 우리가 뽑히는 쪽이 아닌 공고",
     "note": "'위탁운용사 선정' 은 운용사를 뽑는 공고라 입찰 자체가 불가능합니다. "
             "반면 '위탁운용사 선정 정량평가 용역' 은 그 평가를 대행하는 우리 일입니다. "
             "selectee 에 걸리고 delegate 가 하나도 없으면 0점이 됩니다.",
     "fields": [
         F("selectee", "selectee — 피선정 신호", "list", score="걸리면 0점",
           help="단 core 나 market 이 있으면 이 규칙은 쓰지 않습니다."),
         F("delegate", "delegate — 대행 신호 (거부권 해제)", "list", score="거부권 무효",
           help="맨 '평가' 를 넣으면 안 됩니다. '신용평가회사 선정' 의 "
                "기관 이름 속 '평가' 까지 걸려 규칙이 무력해집니다."),
     ]},
    {"group": "점수를 깎는 키워드",
     "fields": [
         F("kill_hard", "kill_hard — 확정 제외", "list", score="-12점씩",
           help="같은 '평가' 라도 우리 업무가 절대 아닌 것. 환경영향평가·회계감사 등"),
         F("kill_soft", "kill_soft — 약한 제외", "list", score="-5점씩",
           help="대체로 아니지만 core 점수가 높으면 살아남습니다."),
     ]},
    {"group": "기관 보강 · PDF 검사",
     "fields": [
         F("inst_extra", "inst_extra — 기관 보강 목록", "list",
           score="위 보강 가점", help="config.yaml 의 institutions 에 없어서 놓치던 기금 운용기관"),
         F("pdf_strong", "pdf_strong — 공고문에서 찾을 키워드", "list",
           score="걸리면 무조건 수집(A)",
           help="자격요건은 여러 업종을 나열하는 형태로 쓰입니다. "
                "3864(집합투자기구평가회사)가 같이 적힌 공고는 우리도 대상입니다."),
     ]},
]

CONFIG_SCHEMA = [
    {"group": "기본",
     "fields": [
         F("service_key", "data.go.kr 인증키", "long",
           help="Encoding / Decoding 어느 쪽을 넣어도 자동 처리합니다. "
                "쓰는 PC마다 따로 발급받으세요 — 호출 한도가 키 단위입니다. "
                "비우면 수집기가 실행되지 않습니다."),
         F("hours", "조회 기간 (시간)", "int",
           help="168 = 7일. 평일 2회·주말 휴무 일정에서는 넓게 잡아야 시간대가 비지 않습니다."),
         F("output_dir", "결과 저장 폴더", "str",
           help="공유폴더는 \\\\서버이름\\... 형태로. Z: 같은 드라이브 문자는 "
                "작업 스케줄러에서 못 찾습니다."),
     ]},
    {"group": "조회 대상",
     "fields": [
         F("targets", "업무구분", "list", help="용역 / 물품 / 공사"),
         F("industry_codes", "무조건 수집할 업종코드", "list",
           help="3865 = 채권평가회사"),
     ]},
    {"group": "대상 기관",
     "note": "이름의 일부만 적으면 됩니다. '공제회' 한 줄로 군인공제회·행정공제회·"
             "교직원공제회가 모두 걸립니다. 여기 있으면 버리는 게 아니라 가점입니다.",
     "fields": [
         F("institutions", "기관 목록 (가점)", "list",
           help="공고기관과 수요기관을 둘 다 검사합니다."),
         F("institutions_low_cut", "검토 문턱을 낮출 기관", "list",
           help="금융공공기관까지 넓히지 마세요. 4개월 14건이 77건이 되고 "
                "대부분 무관한 공고입니다."),
     ]},
    {"group": "메일",
     "fields": [
         F("mail.enabled", "메일 발송", "bool"),
         F("mail.mode", "발송 방식", "enum", options=["outlook", "smtp"],
           help="outlook 은 이 PC의 아웃룩으로. pywin32 필요, 로그온 세션 필요."),
         F("mail.to", "받는 사람", "list"),
         F("mail.cc", "참조", "list"),
         F("mail.draft_only", "보내지 않고 초안만 띄우기", "bool",
           help="테스트 단계에 켜두면 내용을 눈으로 확인할 수 있습니다."),
         F("mail.attach_excel", "공고목록.xlsx 첨부", "bool"),
         F("mail.attach_pdf", "공고문 PDF 첨부", "bool"),
         F("mail.attach_max_mb", "첨부 총 용량 상한 (MB)", "int"),
         F("mail.send_when_empty", "신규 0건일 때도 보내기", "bool",
           help="끄면 받는 쪽에서 '공고가 없었다' 와 '수집기가 죽었다' 를 "
                "구분할 수 없습니다. 켜두는 편이 안전합니다."),
     ]},
    {"group": "메일 문구",
     "note": "{중괄호} 는 보낼 때 값으로 바뀝니다. 없는 이름을 쓰면 그 줄만 "
             "기본 문구로 나가고 화면에 알려줍니다.",
     "fields": [
         F("mail.문구.제목", "제목 (공고 있을 때)", "str",
           help="쓸 수 있는 값: {날짜} {건수} {내역}"),
         F("mail.문구.제목_없음", "제목 (0건일 때)", "str", help="{날짜}"),
         F("mail.문구.첫줄", "본문 첫 줄", "str", help="{시각} {기간}"),
         F("mail.문구.요약", "표 위 요약", "str", help="{건수} {내역}"),
         F("mail.문구.없음", "0건일 때 본문", "str", help="{날짜}"),
         F("mail.문구.안내", "표 아래 안내 (A·B·C 설명)", "text"),
         F("mail.문구.검토안내", "C가 있을 때 덧붙는 안내", "text"),
         F("mail.문구.꼬리말", "맨 아래 작은 글씨", "str", help="{폴더}"),
     ]},
    {"group": "SMTP (발송 방식이 smtp 일 때만)",
     "fields": [
         F("mail.smtp.host", "서버 주소", "str"),
         F("mail.smtp.port", "포트", "int", help="릴레이 25 / 인증 587"),
         F("mail.smtp.tls", "TLS", "bool"),
         F("mail.smtp.user", "계정", "str", help="인증이 필요 없으면 비웁니다."),
         F("mail.smtp.password", "비밀번호", "secret"),
         F("mail.smtp.from", "발신 주소", "str"),
     ]},
]

FILES = {
    "table": {"path": TABLE_PATH, "name": "점수표.yaml", "schema": TABLE_SCHEMA},
    "config": {"path": CONFIG_PATH, "name": "config.yaml", "schema": CONFIG_SCHEMA},
}


def collect(which: str) -> dict:
    """화면에 채워 넣을 현재 값."""
    spec = FILES[which]
    doc = YamlDoc(spec["path"])
    values: dict[str, object] = {}
    for group in spec["schema"]:
        for f in group["fields"]:
            v = doc.get(f["path"])
            if f["type"] == "secret":
                values[f["path"]] = MASK if v else ""
            elif f["type"] == "list":
                values[f["path"]] = [str(x) for x in (v or [])]
            elif f["type"] == "bool":
                values[f["path"]] = bool(v)
            else:
                values[f["path"]] = "" if v is None else v
    return {"values": values, "schema": spec["schema"], "name": spec["name"]}


def dup_report(table: dict) -> list[dict]:
    """여러 목록에 같이 들어 있는 낱말. 일부러 그런 것도 있으니 경고는 아니다."""
    keys = ["core", "market", "target", "action", "selectee", "delegate",
            "kill_hard", "kill_soft"]
    seen: dict[str, list[str]] = {}
    for k in keys:
        for w in table.get(k) or []:
            seen.setdefault(str(w), []).append(k)
    return [{"word": w, "lists": ks} for w, ks in sorted(seen.items())
            if len(ks) > 1]


# ===========================================================================
#  저장 · 채점 · 검증
# ===========================================================================

def apply_changes(which: str, values: dict) -> dict:
    spec = FILES[which]
    doc = YamlDoc(spec["path"])
    for group in spec["schema"]:
        for f in group["fields"]:
            path, kind = f["path"], f["type"]
            if path not in values:
                continue
            v = values[path]
            if kind == "secret":
                if v == MASK:
                    continue            # 고치지 않았다
                doc.set_scalar(path, str(v), f["label"])
            elif kind == "list":
                clean, seen = [], set()
                for x in v:
                    s = str(x).strip()
                    if s and s not in seen:
                        seen.add(s)
                        clean.append(s)
                doc.set_list(path, clean, f["label"])
            elif kind == "int":
                doc.set_scalar(path, int(v), f["label"])
            elif kind == "bool":
                doc.set_scalar(path, bool(v), f["label"])
            else:
                doc.set_scalar(path, str(v), f["label"])

    if not doc.changes:
        return {"ok": True, "changes": [], "backup": None}
    backup = doc.save()
    try:
        shown = str(backup.relative_to(HERE))
    except ValueError:
        shown = backup.name
    return {"ok": True, "changes": doc.changes, "backup": shown}


def score_preview(title: str, org: str, table: dict, config: dict) -> dict:
    """저장하지 않은 상태로 채점한다.

    Scorer 가 점수표·config 경로를 받으므로, 지금 화면의 값을 임시 파일로
    떨어뜨려 그걸 읽히면 된다. 저장 전에 결과를 볼 수 있다.
    """
    from 점수 import Scorer

    with tempfile.TemporaryDirectory() as tmp:
        t = Path(tmp) / "점수표.yaml"
        c = Path(tmp) / "config.yaml"
        t.write_text(yaml.safe_dump(table, allow_unicode=True), encoding="utf-8")
        c.write_text(yaml.safe_dump(config, allow_unicode=True), encoding="utf-8")
        sc = Scorer(table_path=t, config_path=c)
        div, pts, why = sc.verdict(title, org)
        return {"구분": div, "점수": pts, "근거": why,
                "문턱": {"수집": sc.cut_collect, "검토": sc.review_cut(org)}}


def run_verify() -> dict:
    """점수표_검증.py 를 돌린다. 저장된 파일을 읽으므로 저장 뒤에 의미가 있다."""
    if not VERIFY_SCRIPT.exists():
        return {"ok": False, "output": "점수표_검증.py 가 없습니다."}
    env = {**os.environ, "PYTHONUTF8": "1"}
    p = subprocess.run([sys.executable, str(VERIFY_SCRIPT)], cwd=str(HERE),
                       capture_output=True, env=env)
    out = (p.stdout + p.stderr).decode("utf-8", "replace")
    return {"ok": p.returncode == 0, "output": out}


def restore(which: str) -> dict:
    spec = FILES[which]
    stem = spec["path"].stem
    if not BACKUP_DIR.exists():
        return {"ok": False, "message": "되돌릴 백업이 없습니다."}
    baks = sorted(BACKUP_DIR.glob(f"{stem}_*{spec['path'].suffix}"))
    if not baks:
        return {"ok": False, "message": "되돌릴 백업이 없습니다."}
    latest = baks[-1]
    shutil.copy2(latest, spec["path"])
    return {"ok": True, "message": f"{latest.name} 으로 되돌렸습니다."}


# ===========================================================================
#  화면
# ===========================================================================

PAGE = r"""<!doctype html>
<meta charset="utf-8">
<title>설정 편집기</title>
<style>
:root{--bg:#fbfbfa;--fg:#22201d;--dim:#6b6660;--line:#e0ddd7;--card:#fff;
      --accent:#7a5c2e;--warn:#9a3a20;--ok:#2f6b3f;--chip:#f0ede6}
@media(prefers-color-scheme:dark){:root{--bg:#1a1917;--fg:#e8e4dd;--dim:#9a948b;
  --line:#332f2a;--card:#211f1c;--accent:#c9a35f;--warn:#e0836a;--ok:#7fba8c;--chip:#2a2724}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 "Malgun Gothic",system-ui,sans-serif}
header{position:sticky;top:0;z-index:9;background:var(--bg);border-bottom:1px solid var(--line);
  padding:10px 20px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
h1{font-size:15px;margin:0 12px 0 0;font-weight:700}
button{font:inherit;padding:6px 14px;border:1px solid var(--line);background:var(--card);
  color:var(--fg);border-radius:6px;cursor:pointer}
button:hover{border-color:var(--accent)}
button.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
button:disabled{opacity:.45;cursor:default}
.tabs{display:flex;gap:6px;margin-right:auto}
.tab.on{background:var(--chip);border-color:var(--accent);font-weight:700}
main{max-width:940px;margin:0 auto;padding:20px}
.grp{background:var(--card);border:1px solid var(--line);border-radius:10px;
  padding:16px 18px;margin-bottom:16px}
.grp>h2{font-size:14px;margin:0 0 4px}
.note{color:var(--dim);font-size:13px;margin:0 0 14px}
.f{padding:12px 0;border-top:1px solid var(--line)}
.f:first-of-type{border-top:0}
.lab{display:flex;gap:8px;align-items:baseline;flex-wrap:wrap;margin-bottom:6px}
.lab b{font-weight:600}
.score{font-size:12px;color:var(--accent);background:var(--chip);padding:1px 7px;border-radius:10px}
.cnt{font-size:12px;color:var(--dim)}
.help{color:var(--dim);font-size:13px;margin:2px 0 8px}
input[type=text],input[type=number],textarea,select{width:100%;font:inherit;padding:7px 9px;
  border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg)}
textarea{min-height:60px;resize:vertical}
input.mono{font:13px/1.5 Consolas,monospace}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}
.chip{display:inline-flex;align-items:center;gap:6px;background:var(--chip);
  border:1px solid var(--line);border-radius:14px;padding:3px 6px 3px 11px;font-size:14px}
.chip button{border:0;background:none;padding:0 4px;color:var(--dim);font-size:15px;line-height:1}
.chip button:hover{color:var(--warn)}
.chip.new{border-color:var(--accent)}
.row{display:flex;gap:8px}
.row input{flex:1}
label.sw{display:flex;gap:8px;align-items:center;cursor:pointer}
#test{position:sticky;bottom:0;background:var(--card);border:1px solid var(--line);
  border-radius:10px;padding:14px 18px;margin-top:20px}
#test .row{margin-bottom:8px}
#res{font-size:14px}
.v{display:inline-block;padding:2px 10px;border-radius:6px;font-weight:700;margin-right:8px}
.v.수집{background:var(--ok);color:#fff}
.v.검토{background:var(--accent);color:#fff}
.v.제외{background:var(--chip);color:var(--dim)}
pre{white-space:pre-wrap;background:var(--bg);border:1px solid var(--line);
  border-radius:8px;padding:12px;font:13px/1.5 Consolas,monospace;max-height:50vh;overflow:auto}
.msg{padding:10px 14px;border-radius:8px;margin-bottom:14px;font-size:14px}
.msg.ok{background:var(--chip);border:1px solid var(--ok)}
.msg.bad{background:var(--chip);border:1px solid var(--warn);color:var(--warn)}
.dim{color:var(--dim)}
dialog{border:1px solid var(--line);border-radius:10px;background:var(--card);color:var(--fg);
  max-width:760px;width:90%;padding:18px}
dialog::backdrop{background:#0008}
</style>

<header>
  <h1>설정 편집기</h1>
  <div class=tabs>
    <button class="tab on" data-f=table>점수표.yaml</button>
    <button class=tab data-f=config>config.yaml</button>
  </div>
  <span id=dirty class=dim></span>
  <button id=dupbtn>중복 검사</button>
  <button id=verbtn>검증 실행</button>
  <button id=undobtn>되돌리기</button>
  <button id=savebtn class=primary disabled>저장</button>
</header>

<main>
  <div id=msg></div>
  <div id=body></div>

  <div id=test>
    <div class=row>
      <input id=t_title placeholder="공고명을 넣으면 점수가 바로 나옵니다">
      <input id=t_org placeholder="기관명 (예: 국민연금공단)" style="max-width:230px">
    </div>
    <div id=res class=dim>저장하지 않은 지금 화면의 값으로 채점합니다.</div>
  </div>
</main>

<dialog id=dlg><div id=dlgbody></div>
  <div style="text-align:right;margin-top:14px"><button onclick="dlg.close()">닫기</button></div>
</dialog>

<script>
let S={}, cur='table', orig={};
const $=s=>document.querySelector(s);

async function api(p,b){
  const r=await fetch(p,{method:b?'POST':'GET',headers:{'Content-Type':'application/json'},
    body:b?JSON.stringify(b):null});
  return r.json();
}
function clone(o){return JSON.parse(JSON.stringify(o))}
function dirtyCount(){
  let n=0;
  for(const f in S) if(JSON.stringify(S[f].values)!==JSON.stringify(orig[f])) n++;
  return n;
}
function markDirty(){
  const a=JSON.stringify(S[cur].values)!==JSON.stringify(orig[cur]);
  const n=dirtyCount();
  $('#dirty').textContent=n?`고친 파일 ${n}개`:'';
  $('#savebtn').disabled=!n;
  return a;
}

function render(){
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('on',b.dataset.f===cur));
  $('#test').style.display = cur==='table' ? '' : 'none';
  const V=S[cur].values, out=[];
  for(const g of S[cur].schema){
    out.push(`<div class=grp><h2>${esc(g.group)}</h2>`);
    if(g.note) out.push(`<p class=note>${esc(g.note)}</p>`);
    for(const f of g.fields) out.push(field(f,V[f.path]));
    out.push(`</div>`);
  }
  $('#body').innerHTML=out.join('');
  bind();
  markDirty();
}

function esc(s){return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}

function field(f,v){
  const was=orig[cur][f.path];
  const head=`<div class=lab><b>${esc(f.label)}</b>`
    +(f.score?`<span class=score>${esc(f.score)}</span>`:'')
    +(f.type==='list'?`<span class=cnt>${v.length}개</span>`:'')
    +`</div>`+(f.help?`<p class=help>${esc(f.help)}</p>`:'');
  let ctl='';
  if(f.type==='list'){
    const chips=v.map((w,i)=>{
      const isNew=!(was||[]).includes(w);
      return `<span class="chip${isNew?' new':''}"><span>${esc(w)}</span>`
        +`<button data-del="${f.path}" data-i="${i}" title="지우기">×</button></span>`;
    }).join('');
    ctl=`<div class=chips>${chips||'<span class=dim>비어 있음</span>'}</div>`
      +`<div class=row><input type=text data-add="${f.path}" placeholder="낱말을 넣고 Enter — 쉼표로 여러 개">`
      +`<button data-addbtn="${f.path}">추가</button></div>`;
  }else if(f.type==='bool'){
    ctl=`<label class=sw><input type=checkbox data-p="${f.path}" ${v?'checked':''}> ${v?'켜짐':'꺼짐'}</label>`;
  }else if(f.type==='enum'){
    ctl=`<select data-p="${f.path}">`+f.options.map(o=>
      `<option ${o===v?'selected':''}>${esc(o)}</option>`).join('')+`</select>`;
  }else if(f.type==='text'){
    ctl=`<textarea data-p="${f.path}">${esc(v)}</textarea>`;
  }else if(f.type==='int'){
    ctl=`<input type=number data-p="${f.path}" value="${esc(v)}" style="max-width:140px">`;
  }else if(f.type==='long'){
    ctl=`<input type=text class=mono data-p="${f.path}" value="${esc(v)}" spellcheck=false>`;
  }else{
    ctl=`<input type=text data-p="${f.path}" value="${esc(v)}">`;
  }
  return `<div class=f>${head}${ctl}</div>`;
}

function bind(){
  document.querySelectorAll('[data-p]').forEach(el=>{
    el.oninput=()=>{
      const p=el.dataset.p;
      if(el.type==='checkbox') S[cur].values[p]=el.checked;
      else if(el.type==='number'){const n=parseInt(el.value,10);S[cur].values[p]=isNaN(n)?0:n;}
      else S[cur].values[p]=el.value;
      if(el.type==='checkbox'){el.parentNode.lastChild.textContent=' '+(el.checked?'켜짐':'꺼짐');}
      markDirty();
    };
  });
  document.querySelectorAll('[data-del]').forEach(b=>{
    b.onclick=()=>{S[cur].values[b.dataset.del].splice(+b.dataset.i,1);render();};
  });
  document.querySelectorAll('[data-add]').forEach(inp=>{
    inp.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addWords(inp);}};
  });
  document.querySelectorAll('[data-addbtn]').forEach(b=>{
    b.onclick=()=>addWords(document.querySelector(`[data-add="${b.dataset.addbtn}"]`));
  });
}

function addWords(inp){
  const p=inp.dataset.add, list=S[cur].values[p];
  const words=inp.value.split(',').map(s=>s.trim()).filter(Boolean);
  let added=0;
  for(const w of words) if(!list.includes(w)){list.push(w);added++;}
  inp.value='';
  render();
  if(words.length&&!added) note('이미 목록에 있습니다.','bad');
}

function note(t,k){$('#msg').innerHTML=`<div class="msg ${k||'ok'}">${esc(t)}</div>`;
  if(k!=='bad')setTimeout(()=>{if($('#msg').textContent===t)$('#msg').innerHTML='';},6000);}

document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{cur=b.dataset.f;render();});

$('#savebtn').onclick=async()=>{
  $('#savebtn').disabled=true;
  const done=[],fail=[];
  for(const f of ['table','config']){
    if(JSON.stringify(S[f].values)===JSON.stringify(orig[f])) continue;
    const r=await api('/api/save',{file:f,values:S[f].values});
    if(r.ok){orig[f]=clone(S[f].values);done.push(...(r.changes||[]).map(c=>`${S[f].name} — ${c}`));}
    else fail.push(`${S[f].name}: ${r.error}`);
  }
  if(fail.length){note('저장하지 못했습니다. '+fail.join(' / '),'bad');}
  else{
    note('저장했습니다. 백업은 백업/ 폴더에 있습니다.');
    show('저장한 내용',`<pre>${esc(done.join('\n'))}</pre>`
      +`<p class=dim>점수표를 고쳤다면 <b>검증 실행</b> 으로 검증용 공고 8건이 여전히 잡히는지 확인하세요.</p>`);
    await load();
  }
  markDirty();
};

$('#verbtn').onclick=async()=>{
  if(dirtyCount()){note('검증은 저장된 파일을 읽습니다. 먼저 저장하세요.','bad');return;}
  $('#verbtn').disabled=true;$('#verbtn').textContent='검증 중…';
  const r=await api('/api/verify',{});
  $('#verbtn').disabled=false;$('#verbtn').textContent='검증 실행';
  show(r.ok?'검증 통과':'검증 실패 — 점수표를 다시 보세요',`<pre>${esc(r.output)}</pre>`);
};

$('#dupbtn').onclick=async()=>{
  const r=await api('/api/dupes',{table:tableData()});
  const rows=r.dupes.map(d=>`<tr><td style="padding:2px 12px 2px 0"><b>${esc(d.word)}</b></td>`
    +`<td class=dim>${esc(d.lists.join(' · '))}</td></tr>`).join('');
  show('여러 목록에 같이 있는 낱말',
    rows?`<p class=dim>일부러 그렇게 둔 것도 있습니다(공정가치는 core·action 양쪽에 필요). 확인용입니다.</p><table>${rows}</table>`
       :'<p>겹치는 낱말이 없습니다.</p>');
};

$('#undobtn').onclick=async()=>{
  if(!confirm(`${S[cur].name} 을 가장 최근 백업으로 되돌립니다. 지금 화면의 수정은 사라집니다.`))return;
  const r=await api('/api/restore',{file:cur});
  note(r.message,r.ok?'ok':'bad');
  if(r.ok) await load();
};

function tableData(){
  const V=S.table.values,o={};
  for(const g of S.table.schema)for(const f of g.fields)o[f.path]=V[f.path];
  return o;
}
function nest(flat){
  const o={};
  for(const k in flat){
    const ps=k.split('.');let c=o;
    for(let i=0;i<ps.length-1;i++)c=c[ps[i]]??=({});
    c[ps.at(-1)]=flat[k];
  }
  return o;
}

let tmr;
function testScore(){
  clearTimeout(tmr);
  tmr=setTimeout(async()=>{
    const title=$('#t_title').value.trim();
    if(!title){$('#res').className='dim';
      $('#res').textContent='저장하지 않은 지금 화면의 값으로 채점합니다.';return;}
    const r=await api('/api/score',{title,org:$('#t_org').value.trim(),
      table:nest(tableData()),config:nest(configData())});
    if(r.error){$('#res').className='dim';$('#res').textContent=r.error;return;}
    $('#res').className='';
    $('#res').innerHTML=`<span class="v ${esc(r.구분)}">${esc(r.구분)}</span>`
      +`<b>${r.점수}점</b> <span class=dim>— 수집 ${r.문턱.수집}점 / 검토 ${r.문턱.검토}점</span>`
      +(r.근거?`<div class=dim style="margin-top:4px">${esc(r.근거)}</div>`:'');
  },250);
}
function configData(){
  const V=S.config.values,o={};
  for(const g of S.config.schema)for(const f of g.fields){
    if(f.type==='secret'&&V[f.path]===${MASK_JSON}) continue;
    o[f.path]=V[f.path];
  }
  return o;
}
$('#t_title').oninput=testScore;$('#t_org').oninput=testScore;

function show(t,h){$('#dlgbody').innerHTML=`<h2 style="margin:0 0 10px;font-size:15px">${esc(t)}</h2>${h}`;dlg.showModal();}

async function load(){
  for(const f of ['table','config']) S[f]=await api('/api/data?file='+f);
  for(const f of ['table','config']) orig[f]=clone(S[f].values);
  render();
}
addEventListener('beforeunload',e=>{if(dirtyCount()){e.preventDefault();e.returnValue='';}});
load();
</script>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass                      # 콘솔을 조용하게 둔다

    def _send(self, obj, ctype="application/json"):
        body = (obj if isinstance(obj, bytes)
                else json.dumps(obj, ensure_ascii=False).encode("utf-8"))
        self.send_response(200)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.startswith("/api/data"):
            which = "config" if "file=config" in self.path else "table"
            return self._send(collect(which))
        page = PAGE.replace("${MASK_JSON}", json.dumps(MASK, ensure_ascii=False))
        self._send(page.encode("utf-8"), "text/html")

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        req = json.loads(self.rfile.read(n) or b"{}")
        try:
            if self.path == "/api/save":
                return self._send(apply_changes(req["file"], req["values"]))
            if self.path == "/api/score":
                return self._send(score_preview(req["title"], req.get("org", ""),
                                                req["table"], req["config"]))
            if self.path == "/api/verify":
                return self._send(run_verify())
            if self.path == "/api/dupes":
                return self._send({"dupes": dup_report(req["table"])})
            if self.path == "/api/restore":
                return self._send(restore(req["file"]))
        except Exception as exc:                       # noqa: BLE001
            return self._send({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
        self._send({"ok": False, "error": "알 수 없는 요청"})


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main() -> int:
    for p in (CONFIG_PATH, TABLE_PATH):
        if not p.exists():
            sys.exit(f"{p.name} 이 없습니다. 수집기와 같은 폴더에서 실행하세요.")

    port = free_port()
    url = f"http://127.0.0.1:{port}/"
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("=" * 62)
    print(" 설정 편집기")
    print(f"   {url}")
    print("   브라우저가 자동으로 열립니다. 닫으려면 이 창에서 Ctrl+C.")
    print("   고친 내용은 저장을 눌러야 파일에 들어갑니다.")
    print("=" * 62)
    threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n닫았습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
