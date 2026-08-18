@echo off
setlocal
cd /d "%~dp0"

rem 이 파일은 반드시 CP949(ANSI) + CRLF 로 저장해야 한다.
rem UTF-8 로 저장하면 cmd 가 읽는 위치를 잃어 명령이 앞부분부터 잘린 채
rem 실행된다. README.md 의 2-6 절 참고.

rem 작업 스케줄러는 이 파일의 경로만 안다. 안에서 exe 를 부르므로
rem 파이썬을 깔 필요도, 스케줄러 등록을 고칠 필요도 없다.

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

if not exist "로그" md "로그"

rem 화면 출력은 exe 가 알아서 로그\실행.log 에 남긴다.
rem 여기서 잡는 것은 그 기록이 시작되기도 전에 죽은 경우다.
"%~dp0나라장터수집기.exe" 수집 2> "로그\오류.log"

set RC=%ERRORLEVEL%

if not "%RC%"=="0" (
    echo.
    echo ============================================
    echo  오류가 발생했습니다. 종료코드 %RC%
    echo  자세한 내용은 로그\오류.log 를 확인하세요.
    echo ============================================
    rem 스케줄러로 돌면 auto 인자가 붙는다. 아무도 안 보므로 멈추지 않는다.
    if /i not "%~1"=="auto" pause
)

exit /b %RC%
