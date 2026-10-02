@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
if errorlevel 1 goto folderfail
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
py -3 -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 (
  py -3 -u bootstrap.py --check
  goto finish
)
python -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 (
  python -u bootstrap.py --check
  goto finish
)
echo Python 3.10 이상이 필요합니다. Python 설치 후 다시 실행하세요.
echo https://www.python.org/downloads/windows/
echo 설치 화면에서 Add python.exe to PATH를 선택하세요.
popd
pause
exit /b 1
:finish
set "RESULT=%ERRORLEVEL%"
popd
pause
exit /b %RESULT%
:folderfail
echo 폴더에 접근할 수 없습니다. ZIP을 전체 압축 해제 후 실행하세요.
pause
exit /b 1
