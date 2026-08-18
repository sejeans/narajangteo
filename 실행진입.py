# -*- coding: utf-8 -*-
"""exe 하나로 묶을 때의 진입점.

exe 를 여러 개 만들어 서로 찾게 하면 경로가 어긋날 때마다 깨진다.
그래서 실행 파일은 하나만 두고 첫 인자로 역할을 나눈다.

    나라장터수집기.exe              편집기 (더블클릭하면 이것)
    나라장터수집기.exe 수집 [옵션]   수집 (실행.bat / 작업 스케줄러가 부름)
    나라장터수집기.exe 점수표검증    점수표 회귀 검사

편집기가 수집기를 따로 띄울 때도 sys.executable 이 이 exe 자신이라
같은 경로를 그대로 쓴다. 편집기.py 의 child() 를 보라.

소스로 돌릴 때는 이 파일을 거치지 않는다. 편집기.py·수집기.py 를
지금까지처럼 직접 실행하면 된다.
"""
from __future__ import annotations

import sys


def _utf8_출력():
    """화면 출력을 UTF-8 로 고정한다.

    소스로 돌릴 때는 실행.bat 의 PYTHONUTF8=1 이 이 일을 했다. 그런데
    PyInstaller 로 묶으면 그 환경변수가 먹지 않는다. 그대로 두면 출력이
    파이프나 파일로 갈 때 cp949 가 되어, 화살표('—') 같은 글자 하나에
    UnicodeEncodeError 로 통째로 죽는다. 다음 두 경로가 모두 파이프다.

      - 편집기의 '실행' 탭 (수집기를 자식으로 띄우고 출력을 읽는다)
      - 실행.bat 의 2> 로그\\오류.log

    콘솔로 바로 나갈 때는 이미 UTF-8 이라 이 호출이 아무 일도 하지 않는다.
    """
    for s in (sys.stdout, sys.stderr):
        if s is not None:
            try:
                s.reconfigure(encoding="utf-8", errors="replace")
            except (AttributeError, OSError):
                pass


_utf8_출력()


USAGE = """나라장터 공고 수집기

  나라장터수집기.exe             설정 편집기를 브라우저에 엽니다
  나라장터수집기.exe 수집        공고를 수집합니다 (자동 실행이 쓰는 명령)
  나라장터수집기.exe 점수표검증  점수표가 아직 맞는지 검사합니다

수집 뒤에 붙일 수 있는 옵션은 편집기의 '실행' 탭에 정리되어 있습니다.
"""


def 설정_꺼내기() -> None:
    """exe 옆에 설정이 없으면 번들의 기본값을 꺼내 놓는다.

    어느 명령으로 들어오든 필요하므로 여기 한 군데서만 한다.
    이미 있으면 손대지 않으니 업데이트해도 고쳐 둔 설정은 그대로다.
    """
    import 앱경로
    if not 앱경로.FROZEN:
        return
    앱경로.seed("점수표.yaml")
    앱경로.seed("config.yaml", "config.example.yaml")


def 수집(args: list) -> int:
    """수집기를 돌린다. 수집기.py 의 __main__ 블록과 같은 처리를 한다."""
    # 로그 헤더에 옵션이 그대로 찍히도록 서브커맨드는 떼고 넘긴다.
    sys.argv = [sys.argv[0], *args]

    import 수집기
    수집기.enable_log()
    try:
        return 수집기.main()
    except KeyboardInterrupt:
        수집기.log("\n중단했습니다.")
        return 1
    except SystemExit as exc:
        # 수집기는 설정이 틀리면 sys.exit("설명") 으로 끝낸다. 그 설명을
        # 삼키면 스케줄러 로그에 아무것도 안 남아 원인을 못 찾는다.
        if isinstance(exc.code, str):
            수집기.log(f"\n{exc.code}")
            return 1
        return 0 if exc.code in (0, None) else 1
    except Exception as exc:  # noqa: BLE001
        수집기.log(f"\n[예기치 못한 오류] {type(exc).__name__}: {exc}")
        return 1


def 점수표검증() -> int:
    import 점수표_검증
    return 점수표_검증.main()


def 편집기() -> int:
    import 편집기 as mod
    return mod.main()


def 창_붙잡기() -> None:
    """오류 메시지를 읽을 틈을 준다.

    더블클릭으로 띄운 창은 프로그램이 끝나는 순간 닫힌다. 그래서 뜨자마자
    죽으면 '검은 창이 깜빡하고 사라졌다' 로만 보이고 이유를 알 수 없다.

    수집 명령에는 붙이지 않는다. 그쪽은 실행.bat 과 편집기가 부르는데,
    작업 스케줄러 뒤에서 입력을 기다리면 작업이 '실행 중' 인 채로 영영
    남아 다음 회차까지 막는다. 멈춰 세우는 일은 실행.bat 이 auto 인자를
    보고 판단한다.
    """
    try:
        input("\n창을 닫으려면 Enter 를 누르세요...")
    except (EOFError, KeyboardInterrupt, OSError):
        pass


def main() -> int:
    argv = sys.argv[1:]
    cmd = argv[0] if argv else ""

    설정_꺼내기()

    if cmd in ("", "편집기"):
        # 더블클릭으로 들어오는 길이다. 설정 파일이 깨졌거나 포트를 잡지
        # 못하면 여기서 죽는데, 그대로 두면 창이 닫혀 아무것도 못 읽는다.
        try:
            return 편집기()
        except SystemExit as exc:
            if isinstance(exc.code, str):
                print(f"\n{exc.code}")
                창_붙잡기()
                return 1
            raise
        except Exception as exc:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"\n[오류] {type(exc).__name__}: {exc}")
            창_붙잡기()
            return 1
    if cmd == "수집":
        return 수집(argv[1:])
    if cmd == "점수표검증":
        return 점수표검증()
    if cmd in ("--help", "-h", "/?", "도움말"):
        print(USAGE)
        return 0

    print(f"알 수 없는 명령입니다: {cmd}\n")
    print(USAGE)
    창_붙잡기()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
