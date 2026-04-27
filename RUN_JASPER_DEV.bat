@echo off
setlocal

cd /d "%~dp0jasper-service"

if "%REPORTS_API_BASE_URL%"=="" (
  set "REPORTS_API_BASE_URL=http://192.168.2.112:3300"
)

echo Starting jasper-service with REPORTS_API_BASE_URL=%REPORTS_API_BASE_URL%
echo.

call mvnw.cmd spring-boot:run

pause
