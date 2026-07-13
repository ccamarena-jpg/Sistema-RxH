@echo off
REM ============================================================
REM  Deploy manual de Sistema-RxH
REM  Doble clic en este archivo para subir tus cambios a GitHub.
REM  Vercel desplegara automaticamente al detectar el push.
REM ============================================================
cd /d "%~dp0"

echo.
echo === Subiendo cambios a GitHub ===
echo.

git add -A

REM Si no hay nada que subir, salir sin error
git diff --cached --quiet
if %errorlevel%==0 (
  echo No hay cambios que subir.
  echo.
  timeout /t 3 >nul
  exit /b 0
)

for /f "tokens=* usebackq" %%d in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm'"`) do set FECHA=%%d
git commit -m "Actualizacion %FECHA%"
git push origin main

echo.
echo === Listo. Vercel desplegara en unos segundos. ===
echo.
timeout /t 4 >nul
