@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   PAINEL COMPLETO (time)
REM
REM  Painel com TESTE e PRODUCAO e o botao "Resetar ambiente de teste".
REM  Para o time de desenvolvimento/QA. A operadora usa o
REM  "Abrir Painel de Treinamentos.bat" da pasta de cima (so PRODUCAO).
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
REM Tudo abaixo usa caminho completo e o pacote e achado pelo PYTHONPATH.
set "PYTHONPATH=%~dp0..\."
title Automacao de Treinamentos - Abrindo o painel

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "VENV_PYW=%VENV_DIR%\Scripts\pythonw.exe"

echo ==========================================================================
echo   AUTOMACAO DE TREINAMENTOS - PAINEL
echo ==========================================================================
echo.

REM --- 1) Procura o Python instalado no computador --------------------------
set "PY_CMD="
py -3 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD python -c "import sys" >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD goto :sem_python

REM --- 2) Confere / monta o ambiente virtual -------------------------------
if not exist "%VENV_PY%" goto :montar
"%VENV_PY%" -c "import xlwings, dateutil, tkinter" >nul 2>&1 && goto :pronto
echo [INFO] O ambiente existente nao funciona neste computador. Vou refazer.
rmdir /s /q "%VENV_DIR%"

:montar
echo [PREPARANDO] Montando o ambiente. Aguarde de 1 a 2 minutos.
echo (so na primeira vez em cada computador)
echo.
%PY_CMD% -m venv "%VENV_DIR%"
if errorlevel 1 goto :erro_venv
"%VENV_PY%" -m pip install --upgrade pip >nul 2>&1
"%VENV_PY%" -m pip install -r "%~dp0..\requirements.txt"
if errorlevel 1 goto :erro_deps
"%VENV_PY%" -c "import tkinter" >nul 2>&1
if errorlevel 1 goto :sem_tkinter
echo.
echo [OK] Ambiente pronto.

:pronto
REM --- 3) Checagem de import COM console visivel ---------------------------
REM  Se algo estiver quebrado, o erro precisa aparecer aqui - depois que a
REM  janela abre (sem console) nao ha mais onde ler mensagem de import.
"%VENV_PY%" -c "import treinamentos_its.gui" 2>&1
if errorlevel 1 goto :erro_import

echo.
echo Abrindo a janela... (esta caixa preta pode ser fechada)
if exist "%VENV_PYW%" (
    start "" "%VENV_PYW%" -m treinamentos_its.gui
) else (
    start "" "%VENV_PY%" -m treinamentos_its.gui
)
exit /b 0

:sem_python
echo [ERRO] O Python nao foi encontrado neste computador.
echo.
echo   Como resolver:
echo   1. Abra  https://www.python.org/downloads/  e baixe o Python 3.
echo   2. Ao instalar, MARQUE a caixa  "Add Python to PATH".
echo   3. Rode este arquivo de novo.
echo.
pause
exit /b 1

:sem_tkinter
echo.
echo [ERRO] Este Python foi instalado sem o tkinter, que e o que desenha a
echo        janela. Reinstale o Python marcando "tcl/tk and IDLE".
echo        Enquanto isso, use os .bat antigos desta pasta - eles nao precisam
echo        da janela.
echo.
pause
exit /b 1

:erro_import
echo.
echo [ERRO] Nao consegui carregar o painel (mensagem acima).
echo        Use os .bat antigos desta pasta enquanto isso e avise o time.
echo.
pause
exit /b 1

:erro_venv
echo [ERRO] Nao foi possivel criar o ambiente virtual.
pause
exit /b 1

:erro_deps
echo.
echo [ERRO] Nao foi possivel instalar as dependencias.
echo        Verifique a conexao com a internet e rode de novo.
pause
exit /b 1
