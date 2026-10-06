@echo off
REM ==========================================================================
REM  Automacao de Treinamentos - Time de TI   |   CONCLUSAO AVULSA
REM
REM  De DOIS CLIQUES neste arquivo para registrar que UMA pessoa concluiu UM
REM  curso fora do ciclo mensal. O programa pergunta o nome, o codigo do curso
REM  e a data. Marca OK (e limpa o "Planejado") em TODAS as abas de cargo do
REM  "Planos e Macs" do mes em que a pessoa aparece com aquele curso.
REM
REM  Nao altera a VERSAO do curso - isso so muda pela "Matriz - Atualizacao"
REM  no ciclo mensal.
REM ==========================================================================

setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
REM Sem "cd": o cmd nao entra em pasta de servidor (\\servidor\...).
set "RAIZ=%~dp0.."
title Automacao de Treinamentos - Concluir Curso

REM O ambiente Python fica no PC de cada pessoa (nao na pasta do programa):
REM ele e preso ao Python de quem o criou e a pasta pode ser compartilhada.
REM Fora do AppData: o Python da Microsoft Store desvia em silencio o que e
REM gravado em AppData pra outra pasta (o venv nascia num lugar e era
REM procurado em outro - visto no teste do servidor, 06/10/2026).
set "VENV_DIR=%USERPROFILE%\.treinamentos_its\venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo ==========================================================================
echo   REGISTRAR CONCLUSAO DE CURSO
echo ==========================================================================
echo.

set "PY_CMD="
py -3 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD python -c "import sys" >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD goto :sem_python

if not exist "%VENV_PY%" goto :montar
"%VENV_PY%" -c "import xlwings, dateutil" >nul 2>&1 && goto :perguntar
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

:perguntar
echo.
echo Feche todas as janelas do Excel antes de continuar.
echo.
set "NOME="
set "CODIGO="
set "DATA="
set /p "NOME=Nome do colaborador (exatamente como na planilha): "
set /p "CODIGO=Codigo do curso (ex.: DOC-CRG-0078552): "
set /p "DATA=Data de conclusao dd/mm/aaaa (deixe em branco para HOJE): "

if not defined NOME goto :faltou
if not defined CODIGO goto :faltou

echo.
echo --------------------------------------------------------------------------
echo   Nome  : %NOME%
echo   Curso : %CODIGO%
if defined DATA (echo   Data  : %DATA%) else (echo   Data  : hoje)
echo --------------------------------------------------------------------------
choice /c SN /n /m "Confirma? (S/N): "
if errorlevel 2 goto :cancelado

echo.
REM goto em vez de blocos if (...): um ")" dentro de echo - ou no nome da
REM pessoa - fecha o bloco cedo, o cmd aborta o script e a janela some sem
REM mostrar o resultado.
if not defined DATA goto :concluir_hoje
"%VENV_PY%" "%RAIZ%\main.py" concluir "%NOME%" "%CODIGO%" --data %DATA%
set "RC=%errorlevel%"
goto :concluiu
:concluir_hoje
"%VENV_PY%" "%RAIZ%\main.py" concluir "%NOME%" "%CODIGO%"
set "RC=%errorlevel%"
:concluiu

echo.
echo ==========================================================================
if "%RC%"=="0" goto :registrado
if "%RC%"=="3" goto :registro_incompleto
echo   NAO FOI REGISTRADO - codigo %RC%. Verifique se o nome e o codigo estao
echo   escritos exatamente como aparecem na planilha. Nenhuma planilha ficou
echo   alterada pela metade: o programa desfaz sozinho o que tinha mexido.
goto :fim_registro
:registro_incompleto
echo   *** ATENCAO: NAO FOI REGISTRADO E O PROGRAMA NAO CONSEGUIU DESFAZER TUDO ***
echo   Procure acima a linha "RESTAURACAO INCOMPLETA": ela diz, arquivo por
echo   arquivo, de qual copia restaurar. Feche o Excel, faca isso e so entao
echo   tente de novo. Na duvida, chame o time ANTES de mexer.
goto :fim_registro
:registrado
echo   REGISTRADO. Confira a pessoa no "Planos e Macs" do mes: o curso tem que
echo   estar OK em todas as ocorrencias (a quantidade aparece na linha
echo   "Conclusao confirmada" acima).
:fim_registro
echo ==========================================================================
pause
exit /b %RC%

:faltou
echo.
echo [ERRO] Nome e codigo do curso sao obrigatorios.
pause
exit /b 1

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
