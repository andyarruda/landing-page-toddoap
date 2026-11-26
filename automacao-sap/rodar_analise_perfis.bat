@echo off
REM Analise de perfis x GPA - modo manual. Coloque na pasta "entrada":
REM   perfis.xlsx (AGR_USERS/SUIM) e catalogo_gpa.xlsx (obrigatorios)
REM   usuarios.xlsx (USR02) e solicitacoes_gpa.xlsx (opcionais)
cd /d "%~dp0"
set OPC=
if exist entrada\usuarios.xlsx set OPC=%OPC% --usuarios entrada\usuarios.xlsx
if exist entrada\solicitacoes_gpa.xlsx set OPC=%OPC% --solicitacoes entrada\solicitacoes_gpa.xlsx
if exist config_perfis.json set OPC=%OPC% --config config_perfis.json
python analise_perfis.py --perfis entrada\perfis.xlsx --catalogo entrada\catalogo_gpa.xlsx %OPC%
if errorlevel 1 (pause & exit /b)
for /f %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set HOJE=%%d
start "" "saida\%HOJE%\perfis\relatorio_perfis.html"
pause
