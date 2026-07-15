@echo off
chcp 65001 >nul
title 太炅 Lotto Lab PC·모바일 통합 도우미
echo.
echo ============================================
echo   太炅 Lotto Lab PC·모바일 통합 도우미
echo ============================================
echo.
echo 1. PC 업데이트팩 실행
echo 2. 통합프로젝트 폴더 열기
echo 3. 사용방법 열기
echo 4. 종료
echo.
set /p choice=번호를 입력하세요: 

if "%choice%"=="1" (
  start "" "%~dp0PC_업데이트팩\업데이트_실행.bat"
  goto end
)
if "%choice%"=="2" (
  start "" "%~dp0통합프로젝트_v6.0"
  goto end
)
if "%choice%"=="3" (
  start "" notepad "%~dp0사용방법.txt"
  goto end
)
if "%choice%"=="4" goto end

echo 잘못 입력했습니다.
pause

:end
