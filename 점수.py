# -*- coding: utf-8 -*-
"""공고명 기반 점수 계산

점수표.yaml 을 읽어서 공고 제목에 점수를 매긴다.
점수표를 고칠 때는 이 파일이 아니라 점수표.yaml 을 고치고,
고친 뒤 점수표_검증.py 를 돌려 검증용 공고 8건이 여전히 잡히는지 확인한다.

    from 점수 import Scorer
    sc = Scorer()
    점수, 근거 = sc.score("대체투자자산 공정가치평가 용역", "국민연금공단")
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    sys.exit("필요한 패키지를 설치하세요:\n\n    pip install pyyaml\n")

HERE = Path(__file__).resolve().parent
TABLE_PATH = HERE / "점수표.yaml"
CONFIG_PATH = HERE / "config.yaml"

# 제목 비교 전에 지우는 문자들. '공정가치 평가' 와 '공정가치평가',
# '대체투자자산(부동산)' 과 '대체투자자산 부동산' 을 같게 보기 위한 것.
_STRIP = re.compile(r"[\s·,()\[\]「」『』【】‧・:：/]")


def norm(s: str) -> str:
    return _STRIP.sub("", s or "")


class Scorer:
    def __init__(self, table_path: Path | None = None,
                 config_path: Path | None = None):
        tp = table_path or TABLE_PATH
        if not tp.exists():
            sys.exit(f"점수표가 없습니다: {tp}")
        with tp.open(encoding="utf-8") as f:
            t = yaml.safe_load(f) or {}

        self.cut_collect = int(t.get("cut_collect", 13))
        self.cut_review = int(t.get("cut_review", 7))
        self.cut_review_low = int(t.get("cut_review_low", self.cut_review))
        self.inst_bonus = int(t.get("inst_bonus", 3))
        self.inst_extra_bonus = int(t.get("inst_extra_bonus", 2))

        # 미리 정규화해 둔다. 공고 3만 건을 돌리므로 매번 하면 느리다.
        def prep(key):
            return [(w, norm(w)) for w in (t.get(key) or []) if str(w).strip()]

        self.core = prep("core")
        self.market = prep("market")
        self.target = prep("target")
        self.action = prep("action")
        self.selectee = prep("selectee")
        self.delegate = prep("delegate")
        self.kill_hard = prep("kill_hard")
        self.kill_soft = prep("kill_soft")
        self.inst_extra = [str(w) for w in (t.get("inst_extra") or [])]
        self.pdf_strong = [str(w) for w in (t.get("pdf_strong") or [])]

        cp = config_path or CONFIG_PATH
        self.institutions = []
        self.institutions_low_cut = []
        if cp.exists():
            with cp.open(encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            self.institutions = [str(v).strip()
                                 for v in (cfg.get("institutions") or [])
                                 if str(v).strip()]
            self.institutions_low_cut = [
                str(v).strip() for v in (cfg.get("institutions_low_cut") or [])
                if str(v).strip()]

    # ------------------------------------------------------------------
    def score(self, title: str, org: str = "") -> tuple[int, list[str]]:
        """(점수, 근거 목록) 반환."""
        t = norm(title)
        hits: list[str] = []
        pts = 0

        co = [w for w, n in self.core if n in t]
        mk = [w for w, n in self.market if n in t]
        tg = [w for w, n in self.target if n in t]
        ac = [w for w, n in self.action if n in t]

        # 거부권 — 우리가 '뽑히는 쪽'이 아닌 공고.
        # core/market 이 있으면 평가 용역임이 확실하므로 규칙을 쓰지 않는다.
        if not co and not mk:
            sel = [w for w, n in self.selectee if n in t]
            if sel and not any(n in t for _, n in self.delegate):
                return 0, [f"피선정측(입찰불가): {sel[0]}"]

        if co:
            pts += 8 * min(len(co), 3)
            hits.append("확정: " + ", ".join(co[:3]))
        if mk:
            pts += 8
            hits.append("채권시장: " + ", ".join(mk[:2]))
        if tg:
            pts += 2 * min(len(tg), 2)
            hits.append("대상: " + ", ".join(tg[:2]))
        if ac:
            pts += 2 * min(len(ac), 2)
            hits.append("행위: " + ", ".join(ac[:2]))
        if tg and ac:
            pts += 4
            hits.append("조합+4")

        # 확정어도 조합도 없으면, 단어 몇 개 걸린 것만으로 올라오지 못하게 자른다.
        # 이게 없으면 '자산' '투자' '평가' 가 흩어져 있는 공고가 전부 올라온다.
        if not co and not mk and not (tg and ac):
            pts = min(pts, 3)

        for i in self.institutions:
            if i in org:
                pts += self.inst_bonus
                hits.append(f"기관+{self.inst_bonus}")
                break
        else:
            for i in self.inst_extra:
                if i in org:
                    pts += self.inst_extra_bonus
                    hits.append(f"기관보강+{self.inst_extra_bonus}")
                    break

        hard = [w for w, n in self.kill_hard if n in t]
        soft = [w for w, n in self.kill_soft if n in t]
        if hard:
            pts -= 12 * len(hard)
            hits.append("확정킬: " + ", ".join(hard[:2]))
        if soft:
            pts -= 5 * len(soft)
            hits.append("약한킬: " + ", ".join(soft[:2]))

        return pts, hits

    # ------------------------------------------------------------------
    def review_cut(self, org: str = "") -> int:
        """검토(C) 문턱. 연기금·공제회 공고는 더 낮게 본다.

        이 기관들은 '자산운용 컨설팅' 처럼 확정어 없이 5~6점에 머무는 공고가
        실제로 우리 일인 경우가 있다. 범위를 넓히면 무관한 공고가 쏟아지므로
        config.yaml 의 institutions_low_cut 에 적힌 곳에만 적용한다.
        """
        if self.institutions_low_cut and any(i in org
                                             for i in self.institutions_low_cut):
            return min(self.cut_review, self.cut_review_low)
        return self.cut_review

    def verdict(self, title: str, org: str = "") -> tuple[str, int, str]:
        """(구분, 점수, 근거문자열) 반환. 구분은 수집 / 검토 / 제외."""
        pts, hits = self.score(title, org)
        if pts >= self.cut_collect:
            div = "수집"
        elif pts >= self.review_cut(org):
            div = "검토"
        else:
            div = "제외"
        return div, pts, " · ".join(hits)

    def pdf_hit(self, text: str) -> list[str]:
        """PDF 본문에서 자격요건 키워드를 찾는다. 3차 단계에서 쓴다."""
        flat = norm(text)
        return [w for w in self.pdf_strong if norm(w) in flat]


if __name__ == "__main__":
    sc = Scorer()
    if len(sys.argv) > 1:
        title = sys.argv[1]
        org = sys.argv[2] if len(sys.argv) > 2 else ""
        div, pts, why = sc.verdict(title, org)
        print(f"{div}  {pts}점   {title}")
        print(f"      {why}")
    else:
        print(f"점수표 로드 완료 — core {len(sc.core)} · market {len(sc.market)} · "
              f"target {len(sc.target)} · action {len(sc.action)}")
        print(f"컷: 수집 {sc.cut_collect}점 / 검토 {sc.cut_review}점")
        print('\n사용법:  python 점수.py "공고명" "기관명"')
