@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   AMBIENTE DE TESTE
REM
REM  De DOIS CLIQUES neste arquivo para rodar o ciclo mensal SEM RISCO:
REM  ele trabalha em cima de COPIAS das planilhas, guardadas na pasta
REM  "_ambiente-teste" (aqui dentro de "_Automacao"). Os arquivos reais em
REM  "Planos e Macs" NAO sao tocados.
REM
REM  Na primeira vez ele copia as planilhas reais para o sandbox. Nas proximas
REM  reaproveita as copias - da pra rodar o mesmo mes varias vezes seguidas.
REM  Para comecar do zero, rode antes o "Resetar Ambiente de Teste.bat".
REM
REM  Para testar um mes especifico:
REM    "Rodar em Ambiente de Teste.bat" --mes 9 --ano 2026   (mes dos POPs)
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
set "RAIZ=%~dp0.."
title Automacao de Treinamentos - Ambiente de Teste

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo ==========================================================================
echo   AUTOMACAO DE TREINAMENTOS - AMBIENTE DE TESTE  (nao toca nos arquivos reais)
echo ==========================================================================
echo.

REM --- 1) Procura o Python instalado no computador --------------------------
set "PY_CMD="
py -3 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD python -c "import sys" >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD goto :sem_python

REM --- 2) Confere / monta o ambiente virtual -------------------------------
if not exist "%VENV_PY%" goto :montar
"%VENV_PY%" -c "import xlwings, dateutil" >nul 2>&1 && goto :pronto
echo [INFO] O ambiente existente nao funciona neste computador. Vou refazer.
rmdir /s /q "%VENV_DIR%"

:montar
echo [PREPARANDO] Montando o ambiente. Aguarde de 1 a 2 minutos.
echo.
%PY_CMD% -m venv "%VENV_DIR%"
if errorlevel 1 goto :erro_venv
"%VENV_PY%" -m pip install --upgrade pip >nul 2>&1
"%VENV_PY%" -m pip install -r "%RAIZ%\requirements.txt"
if errorlevel 1 goto :erro_deps
echo.
echo [OK] Ambiente pronto.

:pronto
echo.
echo --------------------------------------------------------------------------
echo   Feche TODAS as janelas do Excel antes de continuar.
echo   O resultado sai nas copias dentro de "_ambiente-teste".
echo --------------------------------------------------------------------------
echo.
pause
echo.

"%VENV_PY%" "%RAIZ%\main.py" sync --teste %*
set "RC=%errorlevel%"

echo.
echo ==========================================================================
if not "%RC%"=="0" goto :com_erro
echo   TERMINOU SEM ERROS. Confira as planilhas dentro de "_ambiente-teste".
goto :fim

:com_erro
echo   TERMINOU COM ERRO (codigo %RC%). Leia as mensagens acima.

:fim
echo ==========================================================================
pause
exit /b %RC%

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
