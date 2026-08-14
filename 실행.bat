@echo off
setlocal
cd /d "%~dp0"

rem 이 파일은 반드시 CP949(ANSI) + CRLF 로 저장해야 한다.
rem UTF-8 로 저장하거나 chcp 로 코드페이지를 바꾸면
rem cmd 가 파일 읽는 위치를 잃어버려 명령이 잘려서 실행된다.
rem (python 수집기.py 가 thon 수집기.py 로 잘려 실행된 적이 있다.)

rem 로그 파일에 한글이 깨지지 않게 한다.
set PYTHONUTF8=1

set "PY=C:\Users\user\AppData\Local\Python\pythoncore-3.14-64\python.exe"
if not exist "%PY%" set "PY=python"

if not exist "로그" md "로그"

rem 화면 출력은 수집기.py 가 스스로 로그\실행.log 에 남긴다.
rem 여기서 같은 파일을 리다이렉트로 잡으면 cmd 가 파일을 잠가서
rem 파이썬이 기록을 못 한다. 그래서 오류 출력은 다른 파일로 보낸다.
rem 파이썬이 아예 못 뜨는 경우가 여기에 남는다. 정상이면 빈 파일이다.
"%PY%" "수집기.py" 2> "로그\오류.log"

set RC=%ERRORLEVEL%

if not "%RC%"=="0" (
    echo.
    echo ============================================
    echo  오류가 발생했습니다. 종료코드 %RC%
    echo  자세한 내용은 로그\실행.log 를 확인하세요.
    echo ============================================
    rem 스케줄러 실행 때는 auto 인자가 붙는다. 아무도 못 누르므로 멈추지 않는다.
    if /i not "%~1"=="auto" pause
)

exit /b %RC%
