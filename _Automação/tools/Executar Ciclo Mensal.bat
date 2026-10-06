@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   CICLO MENSAL
REM
REM  De DOIS CLIQUES neste arquivo para rodar o fechamento do mes:
REM    1) cria o "Planos e Macs" do mes novo a partir do mes anterior
REM    2) aplica os versionamentos da "Matriz - Atualizacao" nas abas de cargo
REM    3) escreve o resumo no rodape das abas e atualiza a Matriz
REM  O "Planos e Macs" do mes anterior NUNCA e alterado.
REM
REM  Na primeira vez ele prepara o ambiente sozinho (1 a 2 minutos).
REM  Nas proximas vezes abre direto. Nao e necessario editar este arquivo.
REM
REM  Para um mes especifico, informe o mes DOS POPS (= do arquivo criado):
REM    "Executar Ciclo Mensal.bat" --mes 9 --ano 2026   (POPs de Setembro:
REM    o Planos e Macs de Agosto vira o de Setembro)
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
set "RAIZ=%~dp0.."
title Automacao de Treinamentos - Ciclo Mensal

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo ==========================================================================
echo   AUTOMACAO DE TREINAMENTOS - CICLO MENSAL
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
echo (so na primeira vez em cada computador)
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
echo   ANTES DE CONTINUAR:
echo   - Rode antes o "Conferir Antes de Rodar.bat" e leia os avisos.
echo   - Feche TODAS as janelas do Excel.
echo   - Confirme que a planilha "Planos e Macs" do mes ANTERIOR esta na
echo     pasta "Planos e Macs" do projeto (ao lado da pasta "_Automacao").
echo --------------------------------------------------------------------------
echo.
pause
echo.

"%VENV_PY%" "%RAIZ%\main.py" sync %*
set "RC=%errorlevel%"

echo.
echo ==========================================================================
REM goto em vez de bloco if (...): um ")" dentro de echo fecha o bloco cedo.
if "%RC%"=="0" goto :terminou_ok
if "%RC%"=="3" goto :restauracao_incompleta
echo   TERMINOU COM ERRO - codigo %RC%. Leia as mensagens acima, corrija e
echo   rode de novo.
echo.
echo   O programa DESFEZ sozinho o que tinha alterado neste ciclo: o "Planos e
echo   Macs" do mes novo voltou ao estado de antes (procure "RESTAURADO" acima).
goto :terminou
:restauracao_incompleta
echo   *** ATENCAO: O CICLO FALHOU E O PROGRAMA NAO CONSEGUIU DESFAZER TUDO ***
echo   Procure acima a linha "RESTAURACAO INCOMPLETA": ela diz, arquivo por
echo   arquivo, de qual copia restaurar (ou qual apagar). Feche o Excel, faca
echo   isso e so entao rode de novo. Na duvida, chame o time ANTES de mexer.
goto :terminou
:terminou_ok
echo   TERMINOU SEM ERROS. Leia os avisos (WARNING) acima, principalmente os
echo   ultimos, e confira o "Planos e Macs" do mes e o resumo no rodape das abas.
:terminou
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
