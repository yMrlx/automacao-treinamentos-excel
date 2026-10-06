@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   RESETAR AMBIENTE DE TESTE
REM
REM  De DOIS CLIQUES neste arquivo para APAGAR a pasta "_ambiente-teste" e
REM  recriar as copias das planilhas reais do zero. Use quando quiser rodar o
REM  mesmo mes de novo comecando limpo. Os arquivos reais NAO sao tocados.
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
set "RAIZ=%~dp0.."
title Automacao de Treinamentos - Resetar Ambiente de Teste

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo ==========================================================================
echo   RESETAR AMBIENTE DE TESTE
echo ==========================================================================
echo.

set "PY_CMD="
py -3 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD python -c "import sys" >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD goto :sem_python

if not exist "%VENV_PY%" goto :montar
"%VENV_PY%" -c "import xlwings, dateutil" >nul 2>&1 && goto :rodar
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
echo [OK] Ambiente pronto.
echo.

:rodar
echo Isso vai apagar a pasta "_ambiente-teste" e copiar as planilhas reais de novo.
echo.
choice /c SN /n /m "Confirma? (S/N): "
if errorlevel 2 goto :cancelado

"%VENV_PY%" "%RAIZ%\main.py" resetar-teste
set "RC=%errorlevel%"

echo.
echo ==========================================================================
if not "%RC%"=="0" goto :com_erro
echo   PRONTO. Sandbox "_ambiente-teste" recriado do zero.
goto :fim

:com_erro
echo   TERMINOU COM ERRO (codigo %RC%). Leia as mensagens acima.

:fim
echo ==========================================================================
pause
exit /b %RC%

:cancelado
echo.
echo Cancelado. Nada foi alterado.
pause
exit /b 0

:sem_python
echo [ERRO] O Python nao foi encontrado neste computador.
echo   1. Abra  https://www.python.org/downloads/  e baixe o Python 3.
echo   2. Ao instalar, MARQUE a caixa  "Add Python to PATH".
echo   3. Rode este arquivo de novo.
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
