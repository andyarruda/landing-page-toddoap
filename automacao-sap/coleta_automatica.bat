@echo off
REM Coleta automatica no SAP + analise MM x EWM
REM Pre-requisito: SAP Logon aberto e logado no sistema.
cd /d "%~dp0"
python coletor_sap.py coletar
pause
