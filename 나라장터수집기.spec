# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 빌드 설정.

    빌드.bat        을 더블클릭하거나
    pyinstaller 나라장터수집기.spec --noconfirm

결과는 dist\나라장터 수집기\ 에 나온다. 그 폴더를 통째로 담당자 PC 에
복사하면 파이썬 설치 없이 돈다.

onedir(폴더째) 로 묶는다. onefile 로 하면 실행할 때마다 임시 폴더에
수십 MB 를 풀어야 해서 시작이 느리고, 백신이 더 자주 잡는다.
"""

# 담당자가 편집기로 고치는 파일이므로 번들에는 '기본값' 으로만 넣는다.
# 첫 실행 때 앱경로.seed() 가 exe 옆으로 꺼내고, 그 뒤로는 exe 옆의
# 것만 읽는다. 업데이트해도 고쳐 둔 설정이 덮이지 않는다.
datas = [
    ("config.example.yaml", "."),
    ("점수표.yaml", "."),
]

hiddenimports = [
    # 실행진입.py 가 함수 안에서만 import 하는 것들. 자동 분석이 놓쳐도
    # 들어가도록 못을 박아 둔다.
    "앱경로",
    "수집기",
    "편집기",
    "점수",
    "점수표_검증",
    # 아웃룩 발송용. 설치되어 있지 않으면 수집기가 알아서 건너뛴다.
    "win32com.client",
    "win32timezone",
]

# 쓰지 않는 큰 덩어리들. 빼면 배포 폴더가 눈에 띄게 줄어든다.
excludes = [
    "tkinter", "unittest", "pydoc", "doctest", "test",
    "numpy", "pandas", "matplotlib", "PIL",
]

a = Analysis(
    ["실행진입.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="나라장터수집기",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX 압축은 백신 오탐을 늘린다. 사내 PC 에서는 끈다.
    # 콘솔을 남긴다. 편집기를 띄운 창이 곧 '닫는 방법'(Ctrl+C)이고,
    # 창이 없으면 표준출력이 없어져 자식 프로세스 로그 수집이 꼬인다.
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="나라장터 수집기",
)
