@echo off
setlocal
cd /d "%~dp0"

rem 이 파일은 반드시 CP949(ANSI) + CRLF 로 저장해야 한다. 실행.bat 과 같은 이유다.

rem 이 배치는 개발용이다. 담당자 PC 가 아니라 소스가 있는 이 PC 에서 돌린다.
rem 결과인 dist\나라장터 수집기\ 폴더만 담당자에게 넘기면 된다.

set "PY=C:\Users\user\AppData\Local\Python\pythoncore-3.14-64\python.exe"
if not exist "%PY%" set "PY=python"

echo [1/3] 빌드에 필요한 패키지를 확인합니다.
"%PY%" -m pip install --quiet --upgrade pyinstaller requests pyyaml openpyxl pypdf pywin32
if errorlevel 1 goto fail

echo.
echo [2/3] exe 로 묶습니다. 처음에는 1~2분 걸립니다.
"%PY%" -m PyInstaller "나라장터수집기.spec" --noconfirm
if errorlevel 1 goto fail

echo.
echo [3/3] 배포 폴더에 실행.bat 과 설명서를 넣습니다.
copy /y "배포\실행.bat" "dist\나라장터 수집기\실행.bat" >nul
if errorlevel 1 goto fail
if exist "README.pdf" copy /y "README.pdf" "dist\나라장터 수집기\설명서.pdf" >nul

echo.
echo ============================================
echo  완료했습니다.
echo  dist\나라장터 수집기\ 폴더를 통째로 복사해 넘기세요.
echo ============================================
pause
exit /b 0

:fail
echo.
echo ============================================
echo  빌드에 실패했습니다. 위 메시지를 확인하세요.
echo ============================================
pause
exit /b 1
