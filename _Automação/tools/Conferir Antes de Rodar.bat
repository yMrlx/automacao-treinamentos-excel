@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   CONFERENCIA (so leitura)
REM
REM  De DOIS CLIQUES neste arquivo ANTES de rodar o ciclo mensal.
REM  Ele NAO altera nada. So mostra:
REM    - os versionamentos que o ciclo aplicaria
REM    - o que faria o ciclo PARAR (Matriz ou abas de cargo)
REM    - linhas da Matriz com a coluna F vazia ou com texto (ex.: "VERIFICAR")
REM    - codigos da Matriz que estao no export de "POPs Obsoletos" mais recente
REM
REM  Para conferir um mes especifico:
REM    "Conferir Antes de Rodar.bat" --mes 9 --ano 2026   (mes dos POPs)
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
set "RAIZ=%~dp0.."
title Automacao de Treinamentos - Conferencia

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo ==========================================================================
echo   CONFERENCIA ANTES DE RODAR  (nada e alterado)
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
echo Feche todas as janelas do Excel antes de continuar.
echo.
pause
echo.

"%VENV_PY%" "%RAIZ%\main.py" conferir %*
set "RC=%errorlevel%"

echo.
echo ==========================================================================
REM Antes esta mensagem saia igual com ou sem erro: "Conferencia terminada"
REM em cima de uma checagem que nao aconteceu e confianca falsa.
if "%RC%"=="0" goto :conferiu_ok
echo   A CONFERENCIA ACUSOU PROBLEMA - codigo %RC%. Leia os erros (ERROR)
echo   acima: ou o ciclo vai PARAR por eles, ou alguma checagem nao pode ser
echo   feita. NAO rode o ciclo antes de resolver. Nenhuma planilha foi alterada.
goto :conferiu_fim
:conferiu_ok
echo   Conferencia terminada. Leia os avisos (WARNING) acima antes de rodar
echo   o "Executar Ciclo Mensal.bat". Nenhuma planilha foi alterada.
:conferiu_fim
echo ==========================================================================
pause
exit /b %RC%

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
