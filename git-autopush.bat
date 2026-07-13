@echo off
REM Auto-push silencioso usado por el hook de Claude Code.
REM Sube a GitHub solo si hay cambios. Sin pausas ni ventanas.
cd /d "%~dp0"
git add -A
git diff --cached --quiet && exit /b 0
for /f "tokens=* usebackq" %%d in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm'"`) do set FECHA=%%d
git commit -m "Auto: cambios %FECHA%" >nul 2>&1
git push origin main >nul 2>&1
exit /b 0
