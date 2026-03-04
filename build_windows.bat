@echo off
setlocal

echo ============================================
echo  ExcelMergeSearch - Windows 빌드 스크립트
echo ============================================

rem 현재 스크립트가 있는 디렉터리로 이동
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo.
echo [1/5] 가상환경(venv) 확인 중...

set "VENV_DIR=%SCRIPT_DIR%venv"

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo  - venv가 없어서 새로 생성합니다.
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo  !!! venv 생성에 실패했습니다. Python이 제대로 설치되어 있는지 확인하세요.
        pause
        exit /b 1
    )
) else (
    echo  - 기존 venv를 사용합니다.
)

echo.
echo [2/5] 가상환경 활성화 중...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo  !!! venv 활성화에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo [3/5] 파이썬 패키지 설치/업데이트 중...
pip install --upgrade pip
if exist "%SCRIPT_DIR%requirements.txt" (
    pip install -r "%SCRIPT_DIR%requirements.txt"
) else (
    echo  - requirements.txt 를 찾을 수 없습니다. 이 단계는 건너뜁니다.
)
pip install pyinstaller

echo.
echo [4/5] PyInstaller로 EXE 빌드 중...

set "ICON_PATH=assets\zem-icon.ico"

if exist "%SCRIPT_DIR%%ICON_PATH%" (
    pyinstaller --noconfirm --windowed --onefile ^
      --name "exMerge" ^
      --icon "%ICON_PATH%" ^
      --add-data "%ICON_PATH%;assets" ^
      main.py
) else (
    echo  - 아이콘 파일(%ICON_PATH%)을 찾을 수 없어 아이콘 없이 빌드합니다.
    pyinstaller --noconfirm --windowed --onefile ^
      --name "exMerge" ^
      main.py
)

if errorlevel 1 (
    echo.
    echo  !!! 빌드 중 오류가 발생했습니다. 위의 로그를 확인하세요.
    pause
    exit /b 1
)

echo.
echo [5/5] 빌드 완료!

if exist "%SCRIPT_DIR%dist\exMerge.exe" (
    echo  - dist\ExcelMergeSearch.exe 파일이 생성되었습니다.
    echo.
    echo  dist 폴더를 엽니다...
    start "" "%SCRIPT_DIR%dist"
) else (
    echo  - dist 폴더를 확인해 주세요.
)

echo.
echo 모든 작업이 완료되었습니다. 창을 닫으셔도 됩니다.
pause

endlocal
