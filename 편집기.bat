@echo off
setlocal
cd /d "%~dp0"

rem 이 파일은 반드시 CP949(ANSI) + CRLF 로 저장해야 한다. 실행.bat 과 같은 이유다.
rem UTF-8 로 저장하면 cmd 가 명령을 앞부분부터 잘라 읽는다.

set PYTHONUTF8=1

set "PY=C:\Users\user\AppData\Local\Python\pythoncore-3.14-64\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" "편집기.py"

if not "%ERRORLEVEL%"=="0" (
    echo.
    echo ============================================
    echo  편집기를 띄우지 못했습니다. 위 메시지를 확인하세요.
    echo ============================================
    pause
)
