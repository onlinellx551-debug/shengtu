@echo off
setlocal
set "PROFILE_DIR=%~dp0browser_profiles\taobao_1688_manual"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"
set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%PROFILE_DIR%" --disable-blink-features=AutomationControlled "https://login.taobao.com/member/login.jhtml" "https://login.1688.com/member/signin.htm"
