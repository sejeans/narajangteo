# -*- coding: utf-8 -*-
"""exe 로 묶었을 때와 .py 로 돌릴 때의 경로 차이를 여기서 흡수한다.

PyInstaller 로 묶으면 __file__ 은 `_internal` 안을 가리킨다. 그런데
config.yaml·수집결과·로그는 담당자가 열어봐야 하므로 exe 옆에 있어야 한다.
이 둘을 섞으면 "편집기에서 저장했는데 다음 실행 때 그대로다" 같은,
원인을 찾기 어려운 증상이 난다. 그래서 두 폴더를 이름부터 갈라 놓는다.

    app_dir()     exe 가 놓인 폴더 (소스로 돌리면 소스 폴더)
                  → config.yaml, 점수표.yaml, 수집결과, 로그, 백업
    bundle_dir()  번들에 딸려 들어간 읽기 전용 자료
                  → config.example.yaml 같은 기본값

사용자가 고치는 파일은 번들에 두면 안 된다. 다시 묶기 전까지 못 고치기
때문이다. seed() 가 첫 실행 때 exe 옆으로 꺼내 놓는다.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)

_SRC = Path(__file__).resolve().parent


def app_dir() -> Path:
    """설정·결과·로그가 놓이는 폴더. 담당자가 여는 폴더다."""
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return _SRC


def bundle_dir() -> Path:
    """번들에 들어간 읽기 전용 자료가 있는 폴더."""
    if FROZEN:
        return Path(getattr(sys, "_MEIPASS", _SRC)).resolve()
    return _SRC


def seed(name: str, source: str | None = None) -> Path:
    """exe 옆에 없으면 번들의 기본값을 한 번만 복사한다.

    이미 있으면 손대지 않는다. 담당자가 고쳐 놓은 설정을 프로그램 업데이트가
    덮어쓰는 일은 없어야 한다.
    """
    dst = app_dir() / name
    if dst.exists():
        return dst
    src = bundle_dir() / (source or name)
    if src.exists() and src != dst:
        try:
            shutil.copy2(src, dst)
        except OSError:
            pass
    return dst
