@echo off
echo ================================================
echo   MyFitnessPal - Chrome Logado (com exercicios)
echo ================================================
echo.
echo Iniciando Chrome com modo debug...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9223 --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data MFP" --profile-directory="Default" https://www.myfitnesspal.com/account/login

echo.
echo ================================================
echo   INSTRUCOES:
echo ================================================
echo.
echo 1. Faca login com suas credenciais (MFP):
echo    Email: SEU_EMAIL
echo    Senha: SUA_SENHA
echo.
echo 2. Após logar, acesse seu diário
echo.
echo 3. Depois rode: node mfp_connect_chrome.js --days 7
echo.
echo ================================================
pause
