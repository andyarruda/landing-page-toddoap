@echo off
REM Coleta automatica de perfis no SAP + cruzamento com o catalogo GPA
REM Pre-requisitos: SAP Logon aberto e logado; catalogo exportado do GPA em entrada\catalogo_gpa.xlsx
cd /d "%~dp0"
python coletor_sap.py coletar --fluxo perfis
pause
