@echo off
REM Analise diaria MM x EWM - coloque as exportacoes do dia na pasta "entrada"
REM com os nomes mm.xlsx, ewm.xlsx e (opcional) filas.xlsx, depois de dois cliques aqui.
cd /d "%~dp0"
set FILAS=
if exist entrada\filas.xlsx set FILAS=--filas entrada\filas.xlsx
set CFG=
if exist config.json set CFG=--config config.json
python analise_mm_ewm.py --mm entrada\mm.xlsx --ewm entrada\ewm.xlsx %FILAS% %CFG%
if errorlevel 1 (pause & exit /b)
for /f %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set HOJE=%%d
start "" "saida\%HOJE%\relatorio.html"
pause
