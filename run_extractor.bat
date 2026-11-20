@echo off
chcp 65001 >nul
REM ============================================================
REM OS signal extractor - double-click launcher
REM Finds the .p (or .m) file inside the "data" folder and runs
REM MATLAB directly from there, so there is no need to open MATLAB
REM manually, change folder, or type the function name by hand.
REM (Comments here are kept ASCII-only on purpose: Korean text this
REM  early in the file, before chcp above fully takes effect on
REM  cmd's read-ahead buffer, was observed to get corrupted and
REM  misparsed as stray commands on a Korean-locale Windows system.)
REM ============================================================
setlocal

set "TOOL_DIR=%~dp0"
set "DATA_DIR=%TOOL_DIR%data\"
set "FUNC_NAME="

if not exist "%DATA_DIR%" (
    echo [오류] "data" 폴더를 찾지 못했습니다.
    echo 이 bat 파일은 반드시 "data" 폴더와 같은 위치에 있어야 합니다.
    pause
    exit /b 1
)

REM .p 파일을 최우선으로 찾고, 없으면 .m 파일을 찾는다
for %%f in ("%DATA_DIR%*.p") do if not defined FUNC_NAME set "FUNC_NAME=%%~nf"
if not defined FUNC_NAME for %%f in ("%DATA_DIR%*.m") do if not defined FUNC_NAME set "FUNC_NAME=%%~nf"

if not defined FUNC_NAME (
    echo [오류] "data" 폴더에서 실행할 .p 또는 .m 파일을 찾지 못했습니다.
    pause
    exit /b 1
)

echo "%FUNC_NAME%" 실행 중... MATLAB이 열릴 때까지 잠시 기다려 주세요.
echo (이 창을 닫지 마세요)
echo.

matlab.exe -sd "%DATA_DIR%" -nosplash -r "%FUNC_NAME%"

if errorlevel 9009 (
    echo.
    echo [오류] matlab.exe 를 찾지 못했습니다. 이 컴퓨터에 MATLAB이
    echo 설치되어 있지 않거나, 설치 경로가 시스템 PATH에 등록되어 있지
    echo 않은 것 같습니다.
    echo.
    echo 이 파일을 메모장으로 연 뒤, 아래 줄의 matlab.exe 부분을 실제 설치
    echo 경로(전체 경로^)로 바꿔서 다시 저장하고 실행해 보세요.
    echo   예^) "C:\Program Files\MATLAB\R2023b\bin\matlab.exe" -sd "%%DATA_DIR%%" -nosplash -r "%%FUNC_NAME%%"
    echo.
    pause
)

endlocal
