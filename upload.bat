@echo off
echo Starting upload process...
echo.

chcp 65001 >nul

echo Searching for files...

rem WMIC를 사용하여 로컬 날짜(YYYYMMDD) 가져오기
for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set dt=%%a
set "TODAY=%dt:~0,8%"

set "DOWNLOAD_DIR=C:\Users\hw Shin\Downloads"
set "TARGET_DIR=C:\Users\hw Shin\Desktop\anshim-dashboard\downloads"

rem 폴더가 없으면 생성
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

rem 오늘 날짜 관리리스트 파일 찾기 
set "MGMT_LATEST="
for /f "delims=" %%I in ('dir "%DOWNLOAD_DIR%\안심케어_관리리스트_%TODAY%*.xlsx" /B /O:-D 2^>nul') do (
    set "MGMT_LATEST=%%I"
    goto found_mgmt
)

:found_mgmt
if "%MGMT_LATEST%"=="" (
    echo Today's file not found. Please check Downloads folder. ^(안심케어_관리리스트_%TODAY%^)
    goto error
)

copy /Y "%DOWNLOAD_DIR%\%MGMT_LATEST%" "%TARGET_DIR%\management.xlsx" >nul
echo Management file copied
pause

rem 오늘 날짜 신청변경리스트 파일 찾기
set "REQ_LATEST="
for /f "delims=" %%I in ('dir "%DOWNLOAD_DIR%\안심케어_신청변경리스트_%TODAY%*.xlsx" /B /O:-D 2^>nul') do (
    set "REQ_LATEST=%%I"
    goto found_req
)

:found_req
if "%REQ_LATEST%"=="" (
    echo Today's file not found. Please check Downloads folder. ^(안심케어_신청변경리스트_%TODAY%^)
    goto error
)

copy /Y "%DOWNLOAD_DIR%\%REQ_LATEST%" "%TARGET_DIR%\request.xlsx" >nul
echo Request file copied
pause

cd /d "C:\Users\hw Shin\Desktop\anshim-dashboard"

echo.
echo Uploading to Google Sheets...
node upload-only.js

echo.
echo Done! Check the dashboard.

goto end

:error
echo.
echo Upload process aborted due to missing files.

:end
echo.
echo Press any key to close...
pause > nul
exit /b
