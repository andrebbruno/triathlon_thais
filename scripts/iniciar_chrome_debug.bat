@echo off
echo ================================================
echo   MyFitnessPal - Chrome Debug
echo ================================================
echo.
echo Iniciando Chrome com modo debug...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9223 --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data MFP" --profile-directory="Default" https://www.myfitnesspal.com/food/diary/SEU_USUARIO

echo.
echo Chrome iniciado! Agora:
echo 1. Resolva o captcha se aparecer
echo 2. Faça login se necessário
echo 3. Quando o diario estiver visivel, rode: node mfp_connect_chrome.js
echo.
pause
