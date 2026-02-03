@echo off
set "ROOT=%~dp0.."
set "TARGET=%ROOT%\agents\head\START.md"
set "CODE_EXE=%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe"
if exist "%CODE_EXE%" (
  "%CODE_EXE%" -n "%TARGET%"
) else (
  code -n "%TARGET%"
)
