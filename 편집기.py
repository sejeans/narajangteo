# -*- coding: utf-8 -*-
"""config.yaml · 점수표.yaml 를 브라우저에서 고치는 편집기.

    python 편집기.py            (또는 편집기.bat 더블클릭)

메모장으로 YAML 을 직접 고치면 들여쓰기와 따옴표를 틀리기 쉽고, 고친 결과가
점수에 어떻게 먹히는지는 돌려봐야 알 수 있다. 이 편집기는 값을 고르는 화면을
주고, 저장하기 전에 공고명을 넣어 점수가 어떻게 나오는지 바로 보여준다.
메일 문구도 저장 전에 메일 한 통을 통째로 그려서 보여준다.

화면 위쪽 메뉴는 파일이 아니라 '한 번에 손대는 것' 으로 나눈다. 기본 · 메일 ·
키워드(조회 대상 / 대상 기관 / 점수표) · 수동 실행. 인증키를 넣으러 들어온
사람에게 점수표를 보여줄 이유가 없다. 그래서 한 탭이 두 파일에 걸치기도 한다.

'수동 실행' 탭에서 수집기를 돌릴 수 있다. 수집은 설정을 고치는 것과 성격이 달라
(몇 분 걸리고, API 호출한도를 쓰고, 메일이 실제로 나가고, 수집이력을 바꾼다)
저장 버튼 옆이 아니라 탭을 따로 두고 진행 기록을 그대로 보여준다.

주석을 지우지 않는다.
YAML 전체를 다시 써서 덮는 방식이 아니라, 고칠 줄만 찾아 그 줄만 바꾼다.
두 파일은 절반이 주석이라 덮어쓰면 설명이 통째로 날아간다.
쓰기 직전에 결과를 다시 파싱해서 의도한 값과 같은지 확인하고,
다르면 아예 쓰지 않는다.
"""
from __future__ import annotations

import contextlib
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
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    import yaml
except ImportError:
    sys.exit("필요한 패키지를 설치하세요:\n\n    pip install pyyaml\n")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from 앱경로 import FROZEN, app_dir  # noqa: E402

HERE = app_dir()
if not FROZEN:
    sys.path.insert(0, str(HERE))

CONFIG_PATH = HERE / "config.yaml"
TABLE_PATH = HERE / "점수표.yaml"
BACKUP_DIR = HERE / "백업"
VERIFY_SCRIPT = HERE / "점수표_검증.py"
COLLECTOR = HERE / "수집기.py"


def child(sub: str, args: list | None = None) -> list:
    """수집기·검증기를 따로 띄울 때 쓸 명령줄.

    exe 로 묶으면 sys.executable 이 이 exe 자신이므로 서브커맨드만 붙이면
    된다. .py 로 돌 때는 파이썬에 스크립트 경로를 넘긴다.
    """
    if FROZEN:
        return [sys.executable, sub, *(args or [])]
    script = COLLECTOR if sub == "수집" else VERIFY_SCRIPT
    return [sys.executable, str(script), *(args or [])]

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


# 화면 위쪽 메뉴. 한 파일이 한 탭이 아니라, 한 번에 손대는 것끼리 묶는다.
# 인증키를 고치러 들어온 사람과 키워드를 다듬으러 들어온 사람이 보는 화면이
# 달라야 한다. 키워드 쪽은 항목이 많아 다시 세 갈래로 나눈다.
TABS = [
    {"id": "basic", "label": "기본"},
    {"id": "mail", "label": "메일"},
    {"id": "keyword", "label": "키워드", "subs": [
        {"id": "targets", "label": "조회 대상"},
        {"id": "orgs", "label": "대상 기관"},
        {"id": "score", "label": "점수표"},
    ]},
    {"id": "run", "label": "수동 실행"},
]

# 낱말 조각(칩)의 색. 목록 이름을 외우지 않아도 색만 보고 이 낱말이 점수를
# 올리는 쪽인지 깎는 쪽인지 알 수 있게 한다.
#   up  파랑 — 점수를 올린다            dn  빨강 — 점수를 깎거나 제외한다
#   fr  주황 — 거부권을 푼다            og  초록 — 기관 이름
TONES = [
    {"id": "up", "label": "점수 올림"},
    {"id": "dn", "label": "점수 깎음 · 제외"},
    {"id": "fr", "label": "거부권 해제"},
    {"id": "og", "label": "기관"},
]


TABLE_SCHEMA = [
    {"group": "문턱과 배점", "tab": "keyword", "sub": "score",
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
    {"group": "점수를 올리는 키워드", "tab": "keyword", "sub": "score",
     "note": "공고명에 이 낱말이 있으면 점수가 올라갑니다. "
             "괄호·공백·가운뎃점은 비교 전에 지우므로 '공정가치 평가' 와 "
             "'공정가치평가' 는 같게 봅니다.",
     "fields": [
         F("core", "core — 평가 확정어", "list", score="8점씩 (최대 3개)", tone="up",
           help="이 표현이 있으면 채권평가회사 업무로 봅니다. 아래 거부권도 무효화합니다."),
         F("market", "market — 채권시장 제도·연구", "list", score="8점", tone="up",
           help="'평가' 같은 행위어가 없어도 우리 영역인 주제. "
                "예탁결제원 장외 채권거래 건이 여기 걸립니다."),
         F("target", "target — 평가 대상", "list", score="2점씩 (최대 2개)", tone="up",
           help="모펀드·위탁운용사는 core 가 아니라 여기 둡니다. "
                "core 에 넣으면 '운용사 선정' 공고까지 거부권이 풀립니다."),
         F("action", "action — 평가·산출 행위", "list", score="2점씩 (최대 2개)", tone="up",
           help="target 과 action 이 둘 다 걸리면 +4 가 더 붙습니다."),
     ]},
    {"group": "거부권 — 우리가 뽑히는 쪽이 아닌 공고",
     "tab": "keyword", "sub": "score",
     "note": "'위탁운용사 선정' 은 운용사를 뽑는 공고라 입찰 자체가 불가능합니다. "
             "반면 '위탁운용사 선정 정량평가 용역' 은 그 평가를 대행하는 우리 일입니다. "
             "selectee 에 걸리고 delegate 가 하나도 없으면 0점이 됩니다.",
     "fields": [
         F("selectee", "selectee — 피선정 신호", "list", score="걸리면 0점", tone="dn",
           help="단 core 나 market 이 있으면 이 규칙은 쓰지 않습니다."),
         F("delegate", "delegate — 대행 신호 (거부권 해제)", "list",
           score="거부권 무효", tone="fr",
           help="맨 '평가' 를 넣으면 안 됩니다. '신용평가회사 선정' 의 "
                "기관 이름 속 '평가' 까지 걸려 규칙이 무력해집니다."),
     ]},
    {"group": "점수를 깎는 키워드", "tab": "keyword", "sub": "score",
     "fields": [
         F("kill_hard", "kill_hard — 확정 제외", "list", score="-12점씩", tone="dn",
           help="같은 '평가' 라도 우리 업무가 절대 아닌 것. 환경영향평가·회계감사 등"),
         F("kill_soft", "kill_soft — 약한 제외", "list", score="-5점씩", tone="dn",
           help="대체로 아니지만 core 점수가 높으면 살아남습니다."),
     ]},
    {"group": "기관 보강 · PDF 검사", "tab": "keyword", "sub": "score",
     "fields": [
         F("inst_extra", "inst_extra — 기관 보강 목록", "list", tone="og",
           score="위 보강 가점", help="config.yaml 의 institutions 에 없어서 놓치던 기금 운용기관"),
         F("pdf_strong", "pdf_strong — 공고문에서 찾을 키워드", "list", tone="up",
           score="걸리면 무조건 수집(A)",
           help="자격요건은 여러 업종을 나열하는 형태로 쓰입니다. "
                "3864(집합투자기구평가회사)가 같이 적힌 공고는 우리도 대상입니다."),
     ]},
]

CONFIG_SCHEMA = [
    {"group": "기본", "tab": "basic",
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
    {"group": "조회 대상", "tab": "keyword", "sub": "targets",
     "fields": [
         F("targets", "업무구분", "list",
           help="용역 / 물품 / 공사. ⚠ 물품·공사에는 채권평가회사가 맡을 공고가 "
                "거의 없는 반면 조회량은 몇 배로 늘어 API 일일 호출한도를 넘길 수 "
                "있습니다. 특별한 이유가 없으면 용역만 두세요."),
         F("industry_codes", "무조건 수집할 업종코드", "list", tone="up",
           help="3865 = 채권평가회사"),
     ]},
    {"group": "대상 기관", "tab": "keyword", "sub": "orgs",
     "note": "이름의 일부만 적으면 됩니다. '공제회' 한 줄로 군인공제회·행정공제회·"
             "교직원공제회가 모두 걸립니다. 여기 있으면 버리는 게 아니라 가점입니다.",
     "fields": [
         F("institutions", "기관 목록 (가점)", "list", tone="og",
           help="공고기관과 수요기관을 둘 다 검사합니다."),
         F("institutions_low_cut", "검토 문턱을 낮출 기관", "list", tone="og",
           help="금융공공기관까지 넓히지 마세요. 4개월 14건이 77건이 되고 "
                "대부분 무관한 공고입니다."),
     ]},
    {"group": "메일", "tab": "mail",
     "fields": [
         F("mail.enabled", "메일 발송", "bool"),
         F("mail.mode", "발송 방식", "enum", options=["outlook", "smtp"],
           help="outlook 은 이 PC의 아웃룩으로. pywin32 필요, 로그온 세션 필요."),
         F("mail.to", "받는 사람", "list"),
         F("mail.cc", "참조", "list"),
         F("mail.error_to", "오류 메일 받는 사람", "list",
           help="조회가 실패한 회차에만 갑니다. 업무 담당자에게 API 오류 "
                "메일이 가면 안 읽으므로, 수집기를 손볼 수 있는 사람을 "
                "넣으세요. 비워두면 위 '받는 사람' 으로 갑니다."),
         F("mail.draft_only", "보내지 않고 초안만 띄우기", "bool",
           help="테스트 단계에 켜두면 내용을 눈으로 확인할 수 있습니다."),
         F("mail.attach_excel", "공고목록.xlsx 첨부", "bool"),
         F("mail.attach_pdf", "공고문 PDF 첨부", "bool"),
         F("mail.attach_max_mb", "첨부 총 용량 상한 (MB)", "int"),
         F("mail.send_when_empty", "신규 0건일 때도 보내기", "bool",
           help="끄면 받는 쪽에서 '공고가 없었다' 와 '수집기가 죽었다' 를 "
                "구분할 수 없습니다. 켜두는 편이 안전합니다."),
     ]},
    {"group": "메일 문구", "tab": "mail",
     "note": "{중괄호} 는 보낼 때 값으로 바뀝니다. 항목마다 쓸 수 있는 이름이 "
             "다릅니다. 없는 이름을 쓰면 예외로 죽지 않고 그 줄만 조용히 기본 "
             "문구로 나가므로, 아래 경고를 그때그때 확인하세요.",
     "action": {"id": "mailprev", "label": "메일 미리보기"},
     "fields": [
         F("mail.문구.제목", "제목 (공고 있을 때)", "str",
           vars=["날짜", "건수", "내역"]),
         F("mail.문구.제목_없음", "제목 (0건일 때)", "str", vars=["날짜"]),
         F("mail.문구.제목_일부", "제목 (일부 조회 실패)", "str",
           vars=["날짜", "건수", "내역"]),
         F("mail.문구.제목_실패", "제목 (자동수집 실패, 오류 메일)", "str",
           vars=["날짜"]),
         F("mail.문구.실패알림", "일부 실패했을 때 본문 맨 위 경고", "text",
           vars=["빠짐"]),
         F("mail.문구.첫줄", "본문 첫 줄", "str", vars=["시각", "기간"]),
         F("mail.문구.요약", "표 위 요약", "str", vars=["건수", "내역"]),
         F("mail.문구.없음", "0건일 때 본문", "str", vars=["날짜"]),
         F("mail.문구.안내", "표 아래 안내 (A·B·C 설명)", "text", vars=[]),
         F("mail.문구.검토안내", "C가 있을 때 덧붙는 안내", "text", vars=[]),
         F("mail.문구.꼬리말", "맨 아래 작은 글씨", "str", vars=["폴더"]),
     ]},
    {"group": "SMTP (발송 방식이 smtp 일 때만)", "tab": "mail",
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
    if not FROZEN and not VERIFY_SCRIPT.exists():
        return {"ok": False, "output": "점수표_검증.py 가 없습니다."}
    env = {**os.environ, "PYTHONUTF8": "1"}
    p = subprocess.run(child("점수표검증"), cwd=str(HERE),
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
#  메일 미리보기
# ===========================================================================
#
# 문구를 필드별로 따로 보면 실제 메일이 어떻게 생겼는지 알 수 없다. 안내는 표
# 아래에 붙고 검토안내는 C가 있을 때만 나오며, 0건 메일은 아예 다른 문구를 쓴다.
# 그래서 가짜 공고 몇 건을 만들어 수집기의 mail_subject / mail_html 을 그대로
# 부른다. 여기서 HTML 을 새로 짜면 미리보기와 실제 메일이 갈라져 오히려 해롭다.

_MOD = None          # 수집기 모듈 / 못 읽은 이유(str) / None(아직 안 해봄)


def collector() -> tuple[object, str]:
    """수집기 모듈을 불러온다. (모듈, 오류메시지) 중 한쪽만 찬다.

    수집기는 requests·openpyxl 이 없으면 import 도중 sys.exit 한다. 그것을
    그대로 두면 편집기까지 죽으므로 SystemExit 도 잡는다. 미리보기만 못 쓰고
    나머지 편집 기능은 그대로 돌아야 한다.
    """
    global _MOD
    if _MOD is None:
        try:
            import 수집기 as mod
            _MOD = mod
        except SystemExit as exc:
            _MOD = str(exc).strip() or "수집기.py 를 불러오지 못했습니다."
        except Exception as exc:                       # noqa: BLE001
            _MOD = f"{type(exc).__name__}: {exc}"
    return (None, _MOD) if isinstance(_MOD, str) else (_MOD, "")


@contextlib.contextmanager
def captured(mod):
    """수집기가 log() 로 남기는 [주의] 를 가로채 화면으로 가져온다.

    fill() 은 자리표시자가 틀려도 예외를 내지 않고 기본 문구로 되돌린 뒤
    로그에만 적는다. 그 로그가 미리보기에서 확인해야 할 바로 그 내용이다.
    """
    msgs: list[str] = []
    orig = mod.log
    mod.log = lambda m="": msgs.append(str(m))
    try:
        yield msgs
    finally:
        mod.log = orig


def sample_rows(mod) -> list[list]:
    """미리보기용 가짜 3건.

    A·B·C 를 하나씩 둬야 줄 색과 검토안내까지 다 보인다. 공고종류도
    신규·변경·재공고를 하나씩 두고, 예산이 빈 공고도 한 줄 섞어둔다.
    """
    today = f"{datetime.now():%Y-%m-%d}"
    due = f"{datetime.now() + timedelta(days=7):%Y-%m-%d} 10:00"

    def row(grade, dm, title, score, kw, why, no, kind="신규", budget=""):
        r = [""] * len(mod.HEADERS)
        r[mod.COL_GRADE] = grade
        r[mod.COL_KIND] = kind
        r[mod.COL_BUDGET] = budget
        r[mod.COL_REG] = today
        r[mod.COL_DM] = dm
        r[mod.COL_NT] = "조달청"
        r[mod.COL_TITLE] = title
        r[mod.COL_DUE] = due
        r[mod.COL_SCORE] = score
        r[mod.COL_KW] = kw
        r[mod.COL_WHY] = why
        r[mod.COL_NO] = no
        r[mod.COL_LINK] = "https://www.g2b.go.kr/"
        r[mod.COL_STAMP] = mod.RUN_STAMP
        return r

    return [
        row("A", "경찰공제회", "2026~2027년 대체투자자산 공정가치 평가·검증 용역",
            21, "업종코드 3865", "확정: 공정가치평가 · 기관+3", "20260814001",
            budget=95000000),
        row("B", "새마을금고중앙회",
            "새마을금고중앙회 대체투자자산 공정가치 평가 및 투자성과 모니터링 용역",
            16, "-", "확정: 공정가치평가 · 대상: 대체투자 · 기관+3", "20260814002",
            kind="변경", budget=210000000),
        row("C", "국민연금공단", "기준 포트폴리오 도입을 위한 자산배분 체계 연구용역",
            6, "-", "대상: 자산배분 · 행위: 연구 · 기관+3", "20260814003",
            kind="재공고", budget=""),
    ]


# 미리보기로 볼 수 있는 네 가지. 문구가 갈래마다 통째로 다르다.
PREVIEW_MODES = ("정상", "0건", "일부실패", "오류")


def mail_preview(values: dict, mode: str = "정상") -> dict:
    """저장하지 않은 지금 화면의 문구로 메일 한 통을 통째로 만들어 본다.

    실패 갈래도 여기서 볼 수 있어야 한다. 안 그러면 문구를 고쳐놓고
    진짜 사고가 난 날에 처음 보게 된다.
    """
    mod, err = collector()
    if mod is None:
        return {"ok": False,
                "error": "미리보기는 수집기.py 를 불러올 수 있어야 합니다.\n" + err}

    mode = mode if mode in PREVIEW_MODES else "정상"
    text = {k.rsplit(".", 1)[1]: v for k, v in values.items()
            if k.startswith("mail.문구.")}
    rows = [] if mode == "0건" else sample_rows(mod)
    빠짐 = ("용역 07/01 09:00~07/02 09:00 · 면허제한 06/29~06/30"
           if mode == "일부실패" else "")

    end = datetime.now(mod.KST)
    hours = int(values.get("hours") or 168)
    period = mod.period_phrase(end - timedelta(hours=hours), end)
    root = str(values.get("output_dir") or "./수집결과")

    # 첨부 안내 줄도 설정을 따라간다. 켜둔 줄 모르고 있다가 메일에서 보는 것보다
    # 여기서 보이는 편이 낫다.
    attached: list[Path] = []
    if mode == "오류":
        rows, attached = [], []
        attached.append(Path("공고목록.xlsx"))
    if rows and values.get("mail.attach_pdf"):
        attached += [Path(f"경찰공제회_대체투자자산 공정가치 평가·검증 용역_"
                          f"2026081400{i}.pdf") for i in (1, 2, 3)]

    with captured(mod) as msgs:
        if mode == "오류":
            # 오류 메일은 받는 사람도 내용도 다르다. 관리자에게 가는 진단문이다.
            checks = [mod.check("면허제한", 0, ["06/29 18:00~07/02 18:00"],
                                필수=False),
                      mod.check("용역", 0, ["07/02 06:00~07/02 18:00 (응답 22)"])]
            subject = mod.fill(text, "제목_실패", 날짜=mod.day_stamp(end))
            html, _ = mod.error_body(checks, Path(root),
                                     end - timedelta(hours=hours), end)
        else:
            상태 = mod.일부실패 if 빠짐 else mod.정상
            subject = mod.mail_subject({"mail": {"문구": text}}, rows, end, 상태)
            html = mod.mail_html(rows, period, end, root, attached, text, 빠짐)
    받는이 = ""
    if mode == "오류":
        주소 = (values.get("mail.error_to") or values.get("mail.to") or "")
        if isinstance(주소, list):
            주소 = ", ".join(주소)
        받는이 = str(주소)
    return {"ok": True, "subject": subject, "html": html, "warnings": msgs,
            "받는이": 받는이}


# ===========================================================================
#  수집기 실행
# ===========================================================================
#
# 설정을 고치는 것과 수집기를 돌리는 것은 성격이 다르다. 수집은 몇 분 걸리고,
# API 일일 호출한도를 쓰고, 메일이 실제로 나가고, 수집이력과 엑셀을 바꾼다.
# 그래서 저장 버튼 옆이 아니라 따로 탭을 두고, 진행 기록을 그대로 보여준다.

RUN_LOCK = threading.Lock()
RUN: dict = {"proc": None, "lines": [], "dropped": 0, "code": None,
             "cmd": "", "started": "", "stopped": False, "lock": None}
MAX_LOG_LINES = 5000

ALLOWED_FLAGS = {"--목록만", "--pdf없이", "--메일없이", "--메일만"}
DAYS_RE = re.compile(r"^\d{1,3}(\.\d{1,2})?$")
WHEN_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{1,2}:\d{2})?$")


def check_args(args: list) -> str:
    """브라우저가 보낸 인자를 그대로 믿지 않는다. 통과하면 빈 문자열.

    이 서버는 127.0.0.1 의 임의 포트에 열려 있고, 브라우저에 떠 있는 다른
    페이지도 이 주소로 요청을 보낼 수 있다. 돌릴 수 있는 것은 수집기의 정해진
    옵션뿐이어야 한다.
    """
    i, n_days = 0, 0
    while i < len(args):
        a = args[i]
        if not isinstance(a, str):
            return "인자를 읽지 못했습니다."
        if a in ("--기준", "--저장"):
            v = args[i + 1] if i + 1 < len(args) else None
            if not isinstance(v, str) or not v.strip() or v.startswith("--"):
                return f"{a} 뒤에 값이 없습니다."
            if a == "--기준" and not WHEN_RE.match(v.strip()):
                return f"기준시각 형식이 맞지 않습니다: {v}"
            if a == "--저장" and len(v) > 200:
                return "저장 폴더 경로가 너무 깁니다."
            i += 2
            continue
        if a.startswith("--"):
            if a not in ALLOWED_FLAGS:
                return f"쓸 수 없는 옵션입니다: {a}"
        elif DAYS_RE.match(a):
            n_days += 1
            if n_days > 1:
                return "일수는 하나만 넣을 수 있습니다."
        else:
            return f"쓸 수 없는 인자입니다: {a}"
        i += 1
    return ""


def lock_path(args: list) -> Path | None:
    """수집기가 만드는 _실행중.lock 위치.

    중지를 누르면 수집기가 __exit__ 을 못 타고 죽어 이 파일이 남는다. 그대로
    두면 3시간 동안 '다른 PC에서 이미 실행 중' 으로 막히므로 우리가 치운다.
    """
    raw = None
    for i, a in enumerate(args):
        if a == "--저장" and i + 1 < len(args):
            raw = args[i + 1]
    if raw is None:
        try:
            raw = (yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))
                   or {}).get("output_dir")
        except (OSError, yaml.YAMLError):
            return None
    if not raw:
        return None
    p = Path(str(raw).strip().strip('"'))
    if not p.is_absolute():
        p = HERE / p
    return p / "_실행중.lock"


def clear_lock() -> list[str]:
    """중간에 죽은 수집기가 남긴 _실행중.lock 을 치운다.

    그대로 두면 3시간 동안 '다른 PC에서 이미 실행 중' 으로 막힌다.
    """
    lock = RUN["lock"]
    if lock is None or not lock.exists():
        return []
    try:
        lock.unlink()
        return [f"[중지] 남은 잠금 파일을 지웠습니다: {lock}"]
    except OSError as exc:
        return [f"[중지] 잠금 파일을 지우지 못했습니다. 직접 지우세요: {lock} ({exc})"]


def pump(proc) -> None:
    """수집기가 뱉는 줄을 받아 쌓는다. 화면은 이것을 폴링해서 가져간다."""
    for raw in iter(proc.stdout.readline, b""):
        line = raw.decode("utf-8", "replace").rstrip("\r\n")
        with RUN_LOCK:
            RUN["lines"].append(line)
            over = len(RUN["lines"]) - MAX_LOG_LINES
            if over > 0:
                del RUN["lines"][:over]
                RUN["dropped"] += over
    proc.stdout.close()
    code = proc.wait()

    stopped = RUN["stopped"]
    tail = clear_lock() if stopped else []
    tail.append("[중지했습니다]" if stopped else
                ("[끝났습니다]" if code == 0 else f"[끝났습니다 — 오류 코드 {code}]"))
    with RUN_LOCK:
        RUN["lines"] += tail
        RUN["code"] = code


def start_run(args: list) -> dict:
    if not FROZEN and not COLLECTOR.exists():
        return {"ok": False, "error": "수집기.py 가 없습니다."}
    bad = check_args(args)
    if bad:
        return {"ok": False, "error": bad}

    with RUN_LOCK:
        proc = RUN["proc"]
        if proc is not None and proc.poll() is None:
            return {"ok": False, "error": "이미 실행 중입니다. 끝나면 다시 누르세요."}

    # 자식이 utf-8 로 뱉게 맞춘다. 안 그러면 윈도우 콘솔 코드페이지(949)로
    # 나와 여기서 읽을 때 깨진다.
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    try:
        proc = subprocess.Popen(child("수집", args),
                                cwd=str(HERE), env=env, stdin=subprocess.DEVNULL,
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                creationflags=flags)
    except OSError as exc:
        return {"ok": False, "error": f"실행하지 못했습니다: {exc}"}

    # 화면에 보여줄 명령줄. 손으로 다시 칠 수 있어야 하므로 실제로 돈 것과
    # 같은 모양이어야 한다. exe 배포판에는 파이썬이 없다.
    앞 = "나라장터수집기.exe 수집 " if FROZEN else "python 수집기.py "
    cmd = 앞 + " ".join(
        f'"{a}"' if " " in a else a for a in args)
    with RUN_LOCK:
        RUN.update(proc=proc, lines=[f"$ {cmd}", ""], dropped=0, code=None,
                   cmd=cmd, started=f"{datetime.now():%H:%M:%S}", stopped=False,
                   lock=lock_path(args))
    threading.Thread(target=pump, args=(proc,), daemon=True).start()
    return {"ok": True, "cmd": cmd}


def run_log(want: int) -> dict:
    with RUN_LOCK:
        proc = RUN["proc"]
        start = max(0, want - RUN["dropped"])
        return {"lines": RUN["lines"][start:],
                "next": RUN["dropped"] + len(RUN["lines"]),
                "running": proc is not None and proc.poll() is None,
                "code": RUN["code"], "cmd": RUN["cmd"], "started": RUN["started"]}


def stop_run() -> dict:
    with RUN_LOCK:
        proc = RUN["proc"]
        if proc is None or proc.poll() is not None:
            return {"ok": False, "error": "실행 중이 아닙니다."}
        RUN["stopped"] = True
    proc.terminate()
    return {"ok": True}


# ===========================================================================
#  화면
# ===========================================================================

PAGE = r"""<!doctype html>
<meta charset="utf-8">
<title>설정 편집기</title>
<style>
:root{--bg:#fbfbfa;--fg:#22201d;--dim:#6b6660;--line:#e0ddd7;--card:#fff;
      --accent:#7a5c2e;--warn:#9a3a20;--ok:#2f6b3f;--chip:#f0ede6;
      /* 낱말 조각 색: 파랑 올림 · 빨강 깎음 · 주황 거부권해제 · 초록 기관 */
      --up-fg:#1b4f9c;--up-bg:#e9f1fd;--up-line:#bcd4f5;
      --dn-fg:#a32f22;--dn-bg:#fdeceb;--dn-line:#f2c2bb;
      --fr-fg:#8a5a12;--fr-bg:#fbf2e2;--fr-line:#e9d6ac;
      --og-fg:#256b3d;--og-bg:#e8f5ed;--og-line:#b9dfc7}
@media(prefers-color-scheme:dark){:root{--bg:#1a1917;--fg:#e8e4dd;--dim:#9a948b;
  --line:#332f2a;--card:#211f1c;--accent:#c9a35f;--warn:#e0836a;--ok:#7fba8c;--chip:#2a2724;
  --up-fg:#93b9f2;--up-bg:#17243a;--up-line:#2d4674;
  --dn-fg:#f0968a;--dn-bg:#331a19;--dn-line:#6b3230;
  --fr-fg:#e0b96a;--fr-bg:#32281a;--fr-line:#5f4a26;
  --og-fg:#86c79b;--og-bg:#18291f;--og-line:#2f5a3d}}
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
.subs{display:flex;gap:6px;flex-wrap:wrap;margin:0 0 16px}
.subs button{border-radius:16px;padding:5px 14px}
.subs button.on{background:var(--chip);border-color:var(--accent);font-weight:700}
.subs:empty{display:none}
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
.chip button{border:0;background:none;padding:0 4px;color:inherit;opacity:.55;
  font-size:15px;line-height:1}
.chip button:hover{opacity:1;color:var(--warn)}
/* 이번에 넣은 낱말. 색은 목록의 성격을 나타내므로 여기서 바꾸지 않고 점선으로만. */
.chip.new{border-style:dashed;font-weight:600}

/* 목록의 성격을 색으로. 칩·배지·범례가 같은 규칙을 쓴다.
   .score 와 .chip 뒤에 와야 배경색이 덮인다. */
.t-up{--tfg:var(--up-fg);--tbg:var(--up-bg);--tln:var(--up-line)}
.t-dn{--tfg:var(--dn-fg);--tbg:var(--dn-bg);--tln:var(--dn-line)}
.t-fr{--tfg:var(--fr-fg);--tbg:var(--fr-bg);--tln:var(--fr-line)}
.t-og{--tfg:var(--og-fg);--tbg:var(--og-bg);--tln:var(--og-line)}
.t-up,.t-dn,.t-fr,.t-og{background:var(--tbg);border-color:var(--tln);color:var(--tfg)}
/* 색·테두리는 .t-* 가 준다. 여기서 border 나 color 를 쓰면 선택자가 더 세서
   범례만 회색으로 나온다. */
.legend{display:flex;gap:7px;flex-wrap:wrap;margin:0 0 16px;font-size:13px}
.legend span{border-width:1px;border-style:solid;border-radius:14px;padding:2px 11px}
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
.msg.info{background:var(--chip);border:1px solid var(--line);color:var(--dim)}
.msg:empty{display:none}
.dim{color:var(--dim)}
dialog{border:1px solid var(--line);border-radius:10px;background:var(--card);color:var(--fg);
  max-width:760px;width:90%;padding:18px}
dialog.wide{max-width:1100px}
dialog::backdrop{background:#0008}
.vwarn{color:var(--warn);font-size:13px;margin-top:6px}
.vwarn:empty{display:none}
.rrow{display:flex;gap:7px;align-items:center;flex-wrap:wrap;padding:6px 0}
.rrow input[type=number]{max-width:86px}
.rrow input[type=datetime-local]{max-width:210px;width:auto}
.opt{display:block;padding:5px 0}
.opt .dim{font-size:13px}
#cmdline{margin:0 0 12px;user-select:all}
#runlog{max-height:46vh;min-height:150px;overflow:auto;margin:0}
</style>

<header>
  <h1>설정 편집기</h1>
  <div class=tabs id=tabs></div>
  <span id=dirty class=dim></span>
  <button id=dupbtn>중복 검사</button>
  <button id=verbtn>검증 실행</button>
  <button id=undobtn>되돌리기</button>
  <button id=savebtn class=primary disabled>저장</button>
</header>

<main>
  <div id=msg></div>
  <div class=subs id=subs></div>
  <div id=body></div>

  <div id=runpane style="display:none">
    <div class=grp>
      <h2>조회 범위</h2>
      <p class=note>수집기는 화면 값이 아니라 <b>저장된 파일</b>을 읽습니다.
        고친 것이 있으면 먼저 저장하세요.</p>
      <div class=f>
        <div class=rrow>
          <label class=sw><input type=radio name=rrange value=cfg checked>
            설정된 조회 기간 그대로</label>
          <span class=dim id=cfghours></span>
          <span class=dim>— 매일 자동 실행과 같은 조건</span>
        </div>
        <div class=rrow>
          <label class=sw><input type=radio name=rrange value=days> 최근</label>
          <input type=number id=r_days value=7 min=0.25 step=0.25> 일
          <span class=dim id=daysh></span>
        </div>
        <div class=rrow>
          <label class=sw><input type=radio name=rrange value=when> 기준시각</label>
          <input type=datetime-local id=r_when> 부터 거슬러
          <input type=number id=r_wdays value=0.75 min=0.25 step=0.25> 일
          <span class=dim id=wdaysh></span>
        </div>
        <p class=help>기준시각을 넣으면 그 시각에 돌렸다고 치고 그때까지 올라온
          공고만 봅니다. 지난 회차에 어떤 메일이 나갔을지 다시 만들어 볼 때 씁니다.</p>
      </div>
    </div>

    <div class=grp>
      <h2>옵션</h2>
      <label class="sw opt"><input type=checkbox id=o_list> PDF 저장 안 함
        <span class=dim>(--목록만) 점수표를 다듬을 때. 공고문을 받지 않아
          빠르고 엑셀만 나옵니다</span></label>
      <label class="sw opt"><input type=checkbox id=o_nopdf> PDF 검사 생략
        <span class=dim>(--pdf없이) 훨씬 빠르지만 1·2차만 봅니다.
          공고문에만 업종코드가 적힌 건을 놓칩니다</span></label>
      <label class="sw opt"><input type=checkbox id=o_nomail checked> 메일 보내지 않음
        <span class=dim>(--메일없이) 수집만 하고 발송은 건너뜁니다</span></label>
      <div class=f>
        <div class=lab><b>저장 폴더</b></div>
        <p class=help>비우면 config 의 결과 폴더(<span id=cfgout></span>)에 그대로 씁니다.
          <b>기준시각을 지정할 때는 반드시 다른 폴더로 하세요</b> —
          진짜 수집이력·공고목록.xlsx 에 재현 결과가 섞입니다.</p>
        <input type=text id=r_out placeholder="비우면 config 의 결과 폴더">
      </div>
    </div>

    <div class=grp>
      <div id=runwarn class=msg></div>
      <pre id=cmdline></pre>
      <div class=rrow>
        <button id=runbtn class=primary>실행</button>
        <button id=stopbtn disabled>중지</button>
        <span id=runstat class=dim></span>
      </div>
    </div>

    <div class=grp>
      <h2>실행 기록</h2>
      <pre id=runlog>아직 실행하지 않았습니다.</pre>
    </div>

    <div class=grp>
      <h2>메일만 다시 보내기</h2>
      <p class=note>수집을 하지 않고, 저장 폴더의 공고목록.xlsx 에 있는 마지막
        회차를 그대로 메일로 보냅니다. 발송이 실패했을 때 다시 보내거나 발송
        방식을 시험할 때 씁니다.
        <b>config 의 mail.enabled 가 꺼져 있어도 보냅니다.</b></p>
      <button id=mailonlybtn>마지막 회차 메일만 보내기</button>
    </div>
  </div>

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
const TABS=${TABS_JSON}, TONES=${TONES_JSON};
// cur 는 지금 보고 있는 탭. SUB 는 탭마다 마지막으로 본 하위 메뉴를 기억해
// 탭을 오갈 때 처음으로 되돌아가지 않게 한다.
let S={}, cur=TABS[0].id, SUB={}, orig={}, FMAP={};
const $=s=>document.querySelector(s);
const tabOf=id=>TABS.find(t=>t.id===id);
function subNow(){
  const t=tabOf(cur);
  if(!t||!t.subs) return null;
  return SUB[cur]||(SUB[cur]=t.subs[0].id);
}

async function api(p,b){
  const r=await fetch(p,{method:b?'POST':'GET',headers:{'Content-Type':'application/json'},
    body:b?JSON.stringify(b):null});
  return r.json();
}
function clone(o){return JSON.parse(JSON.stringify(o))}
function dirtyFiles(){
  return ['config','table'].filter(f=>
    S[f]&&JSON.stringify(S[f].values)!==JSON.stringify(orig[f]));
}
function dirtyCount(){return dirtyFiles().length}
function markDirty(){
  const d=dirtyFiles();
  $('#dirty').textContent=d.length?'고친 곳: '+d.map(f=>S[f].name).join(' · '):'';
  $('#savebtn').disabled=!d.length;
  if(cur==='run') runRefresh();
  return d.length;
}

// 지금 탭(과 하위 메뉴)에 속한 그룹을 두 파일에서 모아 온다. 화면의 묶음은
// 파일 경계와 다르다 — '키워드' 에는 config 와 점수표가 같이 들어 있다.
function groupsNow(){
  const s=subNow(), out=[];
  for(const file of ['config','table'])
    for(const g of (S[file]?S[file].schema:[]))
      if(g.tab===cur && (!s || g.sub===s)) out.push([file,g]);
  return out;
}

function render(){
  $('#tabs').innerHTML=TABS.map(t=>
    `<button class="tab${t.id===cur?' on':''}" data-t="${esc(t.id)}">`
    +`${esc(t.label)}</button>`).join('');
  document.querySelectorAll('[data-t]').forEach(b=>
    b.onclick=()=>{cur=b.dataset.t;render();});

  const t=tabOf(cur), s=subNow(), isRun=cur==='run';
  $('#subs').innerHTML=(t.subs||[]).map(x=>
    `<button class="${x.id===s?'on':''}" data-s="${esc(x.id)}">${esc(x.label)}</button>`
    ).join('');
  document.querySelectorAll('[data-s]').forEach(b=>
    b.onclick=()=>{SUB[cur]=b.dataset.s;render();});

  // 점수 시험 칸은 점수표·기관 화면에서만. 조회 대상에는 점수가 걸리지 않는다.
  $('#test').style.display = (cur==='keyword'&&s!=='targets') ? '' : 'none';
  $('#body').style.display = isRun ? 'none' : '';
  $('#runpane').style.display = isRun ? '' : 'none';
  // 중복 검사와 검증은 점수표 이야기라 키워드 탭에서만 뜻이 있다.
  for(const id of ['#dupbtn','#verbtn']) $(id).style.display = cur==='keyword'?'':'none';
  $('#undobtn').style.display = isRun ? 'none' : '';
  if(isRun){ markDirty(); return; }

  const out=[];
  FMAP={};
  // 범례는 이 화면에 실제로 나오는 색만. 안 쓰는 색까지 늘어놓으면
  // 어느 것이 지금 보고 있는 목록의 색인지 되레 헷갈린다.
  const here=groupsNow();
  const tones=TONES.filter(x=>
    here.some(([,g])=>g.fields.some(f=>f.tone===x.id)));
  if(tones.length) out.push(`<div class=legend>`+tones.map(x=>
    `<span class="t-${esc(x.id)}">${esc(x.label)}</span>`).join('')+`</div>`);
  for(const [file,g] of here){
    out.push(`<div class=grp><h2>${esc(g.group)}</h2>`);
    if(g.note) out.push(`<p class=note>${esc(g.note)}</p>`);
    for(const f of g.fields){
      FMAP[f.path]={...f,file};
      out.push(field(file,f,S[file].values[f.path]));
    }
    if(g.action) out.push(`<div class=f><button id="${esc(g.action.id)}">`
      +`${esc(g.action.label)}</button></div>`);
    out.push(`</div>`);
  }
  $('#body').innerHTML=out.join('');
  bind();
  for(const p in FMAP)
    if(FMAP[p].vars) checkVars(FMAP[p], S[FMAP[p].file].values[p]);
  markDirty();
}

// {중괄호} 안의 이름이 그 항목에서 쓸 수 있는 것인지 본다.
// 틀리면 수집기가 예외를 내지 않고 그 줄만 기본 문구로 되돌리므로,
// 메일이 나간 뒤에는 알아채기 어렵다. 여기서 잡는다.
function varProblems(tpl,allowed){
  const t=String(tpl??'').replace(/\{\{|\}\}/g,'');   // {{ 는 글자 그대로의 {
  const bad=[];
  for(const m of t.matchAll(/\{([^{}]*)\}/g)){
    const name=m[1].split(/[:!]/)[0].trim();
    if(!allowed.includes(name)) bad.push('{'+m[1]+'}');
  }
  if(/[{}]/.test(t.replace(/\{[^{}]*\}/g,''))) bad.push('짝이 맞지 않는 중괄호');
  return bad;
}
function checkVars(f,v){
  const el=document.querySelector(`[data-bad="${f.path}"]`);
  if(!el) return;
  const bad=varProblems(v,f.vars);
  el.textContent = bad.length
    ? `⚠ 여기서 쓸 수 없습니다: ${bad.join(', ')} — 이대로 두면 이 줄만 기본 문구로 나갑니다`
    : '';
}

function esc(s){return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}

function field(file,f,v){
  const was=orig[file][f.path];
  const tone=f.tone?' t-'+f.tone:'';
  const head=`<div class=lab><b>${esc(f.label)}</b>`
    +(f.score?`<span class="score${tone}">${esc(f.score)}</span>`:'')
    +(f.type==='list'?`<span class=cnt>${v.length}개</span>`:'')
    +`</div>`+(f.help?`<p class=help>${esc(f.help)}</p>`:'')
    +(f.vars?`<p class=help>쓸 수 있는 값: `
      +(f.vars.length?f.vars.map(x=>'{'+esc(x)+'}').join(' ')
                     :'없음 — 중괄호를 쓰면 기본 문구로 되돌아갑니다')+`</p>`:'');
  const tail=f.vars?`<div class=vwarn data-bad="${esc(f.path)}"></div>`:'';
  let ctl='';
  if(f.type==='list'){
    const chips=v.map((w,i)=>{
      const isNew=!(was||[]).includes(w);
      return `<span class="chip${tone}${isNew?' new':''}"`
        +(isNew?' title="아직 저장하지 않은 낱말"':'')+`><span>${esc(w)}</span>`
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
  return `<div class=f>${head}${ctl}${tail}</div>`;
}

function bind(){
  document.querySelectorAll('[data-p]').forEach(el=>{
    el.oninput=()=>{
      const p=el.dataset.p, V=S[FMAP[p].file].values;
      if(el.type==='checkbox') V[p]=el.checked;
      else if(el.type==='number'){const n=parseInt(el.value,10);V[p]=isNaN(n)?0:n;}
      else V[p]=el.value;
      if(el.type==='checkbox'){el.parentNode.lastChild.textContent=' '+(el.checked?'켜짐':'꺼짐');}
      if(FMAP[p]&&FMAP[p].vars) checkVars(FMAP[p],el.value);
      markDirty();
    };
  });
  const mp=$('#mailprev'); if(mp) mp.onclick=()=>mailPreview('정상');
  document.querySelectorAll('[data-del]').forEach(b=>{
    b.onclick=()=>{
      const p=b.dataset.del;
      S[FMAP[p].file].values[p].splice(+b.dataset.i,1);
      render();
    };
  });
  document.querySelectorAll('[data-add]').forEach(inp=>{
    inp.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addWords(inp);}};
  });
  document.querySelectorAll('[data-addbtn]').forEach(b=>{
    b.onclick=()=>addWords(document.querySelector(`[data-add="${b.dataset.addbtn}"]`));
  });
}

function addWords(inp){
  const p=inp.dataset.add, list=S[FMAP[p].file].values[p];
  const words=inp.value.split(',').map(s=>s.trim()).filter(Boolean);
  let added=0;
  for(const w of words) if(!list.includes(w)){list.push(w);added++;}
  inp.value='';
  render();
  if(words.length&&!added) note('이미 목록에 있습니다.','bad');
}

function note(t,k){$('#msg').innerHTML=`<div class="msg ${k||'ok'}">${esc(t)}</div>`;
  if(k!=='bad')setTimeout(()=>{if($('#msg').textContent===t)$('#msg').innerHTML='';},6000);}

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

// 화면의 탭이 파일 하나에 대응하지 않으므로(키워드 탭은 두 파일을 같이 보여준다)
// 어느 파일을 되돌릴지 여기서 고르게 한다.
$('#undobtn').onclick=()=>{
  show('가장 최근 백업으로 되돌리기',
    `<p class=dim>백업/ 폴더에 있는 가장 최근 파일로 되돌립니다. `
    +`지금 화면에서 고쳐 둔 것은 사라집니다.</p>`
    +['config','table'].map(f=>`<button data-undo="${f}" style="margin-right:8px">`
      +`${esc(S[f].name)} 되돌리기</button>`).join('')
    +`<p class=dim style="margin-top:12px">config.yaml — 기본 · 메일 · 조회 대상 · `
    +`대상 기관 &nbsp; / &nbsp; 점수표.yaml — 키워드 탭의 점수표</p>`);
  document.querySelectorAll('[data-undo]').forEach(b=>b.onclick=async()=>{
    const f=b.dataset.undo;
    if(!confirm(`${S[f].name} 을 가장 최근 백업으로 되돌립니다. 계속할까요?`))return;
    const r=await api('/api/restore',{file:f});
    dlg.close();
    note(r.message,r.ok?'ok':'bad');
    if(r.ok) await load();
  });
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

function show(t,h,wide){
  dlg.classList.toggle('wide',!!wide);
  $('#dlgbody').innerHTML=`<h2 style="margin:0 0 10px;font-size:15px">${esc(t)}</h2>${h}`;
  if(!dlg.open) dlg.showModal();
}

// ---- 메일 미리보기 ------------------------------------------------------
// 저장하지 않은 지금 화면의 문구로 수집기가 메일 한 통을 만들어 돌려준다.
// 0건 메일은 문구가 통째로 달라 따로 봐야 한다.
const MP_MODES=['정상','0건','일부실패','오류'];
const MP_LABELS=['공고 3건 (A·B·C)','신규 0건','일부 조회 실패','자동수집 실패'];
async function mailPreview(mode){
  mode=MP_MODES.includes(mode)?mode:'정상';
  const r=await api('/api/mailpreview',{values:S.config.values,mode:mode});
  if(!r.ok){show('메일 미리보기',`<p class=vwarn>${esc(r.error)}</p>`);return;}
  const warn=(r.warnings||[]).filter(x=>x.trim())
    .map(x=>`<div class="msg bad">${esc(x)}</div>`).join('');
  show('메일 미리보기',
    `<div style="margin-bottom:12px">`
    +MP_MODES.map((m,i)=>`<button class="tab${m===mode?' on':''}" `
      +`data-mp=${i}>${MP_LABELS[i]}</button>`).join(' ')+`</div>`
    +warn
    +(mode==='오류'
      ?`<p class=dim style="margin:0 0 10px">조회가 통째로 실패한 회차에만 나갑니다.`
       +` 받는 사람: <b>${esc(r.받는이||'(비어 있음)')}</b></p>`:'')
    +`<p class=dim style="margin:0 0 3px">제목</p>`
    +`<p style="margin:0 0 14px;font-weight:600">${esc(r.subject)}</p>`
    +`<p class=dim style="margin:0 0 3px">본문</p>`
    +`<iframe id=mpf sandbox style="width:100%;height:50vh;background:#fff;`
    +`border:1px solid var(--line);border-radius:8px"></iframe>`
    +`<p class=dim style="margin:8px 0 0;font-size:13px">가짜 공고로 만든 예시입니다.`
    +` 표에 들어가는 값은 실제 수집 결과로 바뀝니다.</p>`, true);
  $('#mpf').srcdoc='<meta charset="utf-8"><body style="margin:12px;background:#fff">'+r.html;
  document.querySelectorAll('[data-mp]').forEach(b=>
    b.onclick=()=>mailPreview(MP_MODES[+b.dataset.mp]));
}

// ---- 실행 탭 ------------------------------------------------------------
let runFrom=0, runBusy=false, outTouched=false, pollTimer=null;

function rmode(){return document.querySelector('input[name=rrange]:checked').value}
function numOr(v,d){const n=parseFloat(v);return (isNaN(n)||n<=0)?d:n}
function hoursText(v){const n=parseFloat(v);return isNaN(n)?'':`= ${Math.round(n*24)}시간`}

function runArgs(mailOnly){
  const a=[], out=$('#r_out').value.trim();
  if(mailOnly){
    a.push('--메일만');
    if(out) a.push('--저장',out);
    return a;                       // 나머지 옵션은 수집을 안 하므로 뜻이 없다
  }
  const m=rmode();
  if(m==='days') a.push(String(numOr($('#r_days').value,7)));
  if(m==='when'){
    a.push(String(numOr($('#r_wdays').value,0.75)));
    a.push('--기준',($('#r_when').value||'').replace('T',' '));
  }
  if(out) a.push('--저장',out);
  if($('#o_list').checked) a.push('--목록만');
  if($('#o_nopdf').checked) a.push('--pdf없이');
  if($('#o_nomail').checked) a.push('--메일없이');
  return a;
}
function runCmd(mailOnly){
  return 'python 수집기.py '
    + runArgs(mailOnly).map(s=>/\s/.test(s)?`"${s}"`:s).join(' ');
}

// 메일이 실제로 나가는 실행인지. 저장된 값으로 판단한다 — 수집기가 읽는 것이
// 화면 값이 아니라 파일이기 때문이다.
function mailNote(){
  const V=orig.config||{};
  if($('#o_nomail').checked) return ['info','메일은 보내지 않습니다.'];
  if(!V['mail.enabled'])
    return ['info','config 의 mail.enabled 가 꺼져 있어 메일은 나가지 않습니다.'];
  if(V['mail.draft_only'])
    return ['ok','아웃룩에 초안만 띄웁니다. 직접 눌러야 나갑니다.'];
  const to=(V['mail.to']||[]).join(', ');
  return ['bad','⚠ 메일이 실제로 발송됩니다 — '+(V['mail.mode']||'outlook')
    +' · 받는 사람 '+(to||'(비어 있음)')];
}

function runRefresh(){
  const V=orig.config||{};
  const h=parseFloat(V.hours);
  $('#cfghours').textContent = isNaN(h) ? '' : `(${h}시간 = ${+(h/24).toFixed(2)}일)`;
  $('#cfgout').textContent = V.output_dir||'';
  $('#daysh').textContent = hoursText($('#r_days').value);
  $('#wdaysh').textContent = hoursText($('#r_wdays').value);
  $('#cmdline').textContent = runCmd(false);
  const [k,t]=mailNote();
  $('#runwarn').className='msg '+k;
  $('#runwarn').textContent=t;
}

// 기준시각을 넣으면 저장 폴더를 재현용으로 채워 둔다. 비워 둔 채 돌리면
// 진짜 수집이력과 엑셀에 재현 결과가 섞인다.
function autoOut(){
  const v=$('#r_when').value;
  if(rmode()!=='when'||!v||outTouched) return;
  $('#r_out').value='./재현_'+v.replace(/\D/g,'').slice(0,12)
    .replace(/^(\d{8})(\d{4})$/,'$1_$2');
}

function runBlocked(){
  if(dirtyCount()) return '고친 설정을 먼저 저장하세요. 수집기는 저장된 파일을 읽습니다.';
  if(rmode()==='when'&&!$('#r_when').value) return '기준시각을 넣으세요.';
  return '';
}

async function startRun(mailOnly){
  const b=runBlocked();
  if(b){note(b,'bad');return;}
  if(!mailOnly&&rmode()==='when'&&!$('#r_out').value.trim()
     &&!confirm('저장 폴더가 비어 있습니다.\n재현 결과가 진짜 수집이력과 '
       +'공고목록.xlsx 에 섞입니다.\n\n그래도 실행할까요?')) return;
  const r=await api('/api/run',{args:runArgs(mailOnly)});
  if(!r.ok){note(r.error,'bad');return;}
  runFrom=0;
  schedule(0);
}

// 진행 기록은 서버에 쌓이고 화면이 가져간다. 타이머는 하나만 둔다.
function schedule(ms){clearTimeout(pollTimer);pollTimer=setTimeout(poll,ms);}

async function poll(){
  clearTimeout(pollTimer);
  if(runBusy){schedule(300);return;}      // 겹치면 미룬다. 체인을 끊지 않는다
  runBusy=true;
  let r;
  try{ r=await api('/api/runlog?from='+runFrom); }
  catch(e){ runBusy=false; schedule(2000); return; }
  runBusy=false;
  if(r.lines&&r.lines.length){
    const pre=$('#runlog');
    if(runFrom===0) pre.textContent='';   // 안내 문구를 밀어낸다
    const atEnd=pre.scrollTop+pre.clientHeight>=pre.scrollHeight-30;
    pre.textContent+=r.lines.join('\n')+'\n';
    if(atEnd) pre.scrollTop=pre.scrollHeight;
  }
  runFrom=r.next;
  $('#runbtn').disabled=r.running;
  $('#mailonlybtn').disabled=r.running;
  $('#stopbtn').disabled=!r.running;
  $('#runstat').textContent = !r.cmd ? ''
    : r.running ? `실행 중… (${r.started} 시작)`
    : r.code===0 ? `끝났습니다 (${r.started} 시작)`
    : `끝났습니다 — 오류 코드 ${r.code}`;
  if(r.running) schedule(700);
}

$('#runbtn').onclick=()=>startRun(false);
$('#stopbtn').onclick=async()=>{
  if(!confirm('실행을 중지합니다. 그때까지 받은 것은 저장되지 않습니다.'))return;
  const r=await api('/api/runstop',{});
  if(!r.ok) note(r.error,'bad'); else schedule(0);
};
$('#mailonlybtn').onclick=()=>{
  const to=((orig.config||{})['mail.to']||[]).join(', ');
  if(!confirm('수집을 하지 않고, 저장 폴더의 마지막 회차를 그대로 메일로 보냅니다.\n\n'
    +'받는 사람: '+(to||'(비어 있음)')+'\n\n'
    +'config 의 mail.enabled 가 꺼져 있어도 보냅니다. 진행할까요?')) return;
  startRun(true);
};
$('#r_out').oninput=()=>{outTouched=true;runRefresh();};
$('#r_when').onchange=()=>{autoOut();runRefresh();};
document.querySelectorAll('#runpane input').forEach(el=>{
  el.addEventListener('input',runRefresh);
  // 숫자·날짜 칸을 만지면 그 줄의 라디오를 같이 켠다. 값만 고쳐 놓고
  // 왜 안 먹히는지 헤매는 일이 없게.
  const rd=el.closest('.rrow')?.querySelector('input[type=radio]');
  if(rd&&el!==rd) el.addEventListener('focus',()=>{
    if(!rd.checked){rd.checked=true;autoOut();runRefresh();}});
});
document.querySelectorAll('input[name=rrange]').forEach(el=>
  el.addEventListener('change',()=>{autoOut();runRefresh();}));

async function load(){
  for(const f of ['table','config']) S[f]=await api('/api/data?file='+f);
  for(const f of ['table','config']) orig[f]=clone(S[f].values);
  render();
  poll();          // 새로고침 전에 시작한 실행이 있으면 이어서 보여준다
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
        if self.path.startswith("/api/runlog"):
            m = re.search(r"from=(\d+)", self.path)
            return self._send(run_log(int(m.group(1)) if m else 0))
        page = PAGE
        for name, value in (("MASK_JSON", MASK), ("TABS_JSON", TABS),
                            ("TONES_JSON", TONES)):
            page = page.replace("${" + name + "}",
                                json.dumps(value, ensure_ascii=False))
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
            if self.path == "/api/mailpreview":
                return self._send(mail_preview(req["values"],
                                               str(req.get("mode") or "정상")))
            if self.path == "/api/run":
                return self._send(start_run(list(req.get("args") or [])))
            if self.path == "/api/runstop":
                return self._send(stop_run())
        except Exception as exc:                       # noqa: BLE001
            return self._send({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
        self._send({"ok": False, "error": "알 수 없는 요청"})


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main() -> int:
    # exe 로 묶었을 때 설정 파일을 exe 옆으로 꺼내는 일은 실행진입.py 가 한다.
    # 어느 명령으로 들어오든 필요해서 진입점 한 군데로 모아 두었다.
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
    print("   '수동 실행' 탭에서 수집기를 돌릴 수 있습니다. 이 창을 닫으면 같이 멈춥니다.")
    print("=" * 62)
    threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n닫았습니다.")
    finally:
        # 돌고 있는 수집기를 두고 나가면 아무도 읽지 않는 파이프가 가득 차
        # 그대로 멈춰 서고, 잠금 파일이 남아 다음 실행까지 막는다.
        proc = RUN["proc"]
        if proc is not None and proc.poll() is None:
            print("실행 중이던 수집기를 중지합니다...")
            RUN["stopped"] = True
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                proc.kill()
            for line in clear_lock():
                print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
