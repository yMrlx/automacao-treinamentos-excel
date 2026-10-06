"""Ambiente de teste (sandbox) da automação.

Faz uma cópia isolada das planilhas reais em ``_Automação/_ambiente-teste/`` pra
o time rodar o mesmo mês quantas vezes quiser sem risco de bagunçar os
arquivos de produção. A flag ``--teste`` do ``sync``/``conferir`` liga esse
modo; o subcomando ``resetar-teste`` recria as cópias do zero.

Este módulo NÃO importa ``config`` de propósito: os caminhos do sandbox
precisam ser resolvidos ANTES de o ``config`` rodar (o ``main.py`` seta a env
var ``TREINAMENTOS_ITS_BASE`` apontando pra cá antes de importar a cli, porque
os caminhos do ``config`` são resolvidos no import).
"""

from __future__ import annotations

import shutil
import stat
import time
from pathlib import Path

from treinamentos_its.logger import log

# treinamentos_its/ambiente_teste.py -> _Automação -> raiz do projeto
_ESTE_ARQUIVO = Path(__file__).resolve()
PASTA_AUTOMACAO = _ESTE_ARQUIVO.parents[1]
RAIZ_PROJETO = _ESTE_ARQUIVO.parents[2]

PASTA_SANDBOX = PASTA_AUTOMACAO / "_ambiente-teste"

# --- "esta execução é em modo teste?" -- decidido em UM lugar só --------------
#
# Existiam DUAS decisões independentes sobre isso e elas podiam discordar: o
# `main.py` procurava a string exata "--teste" no argv (pra preparar o sandbox
# ANTES de o config resolver os caminhos) e o `argparse` do `cli` decidia de
# novo, aceitando ABREVIAÇÃO de flag. Resultado: `python main.py sync --tes`
# rodava com `args.teste = True` e os caminhos de PRODUÇÃO.
#
# Agora: (1) a regra mora aqui e o `main.py` usa esta função; (2) o `cli` cria
# os parsers com `allow_abbrev=False`, então o argparse só aceita `--teste`
# escrito por inteiro — exatamente o que esta função procura; (3) o `cli` ainda
# confere, antes de executar qualquer coisa, se a base realmente é o sandbox
# (`modo_teste_ativo`) e aborta se não for. Pra reintroduzir o furo seria
# preciso desfazer os três.
FLAG_TESTE = "--teste"
# "concluir" entrou aqui em 2026-09-21, junto com a GUI: a janela tem um botão
# de registrar conclusão e, sem a flag, esse botão escreveria nas planilhas
# REAIS mesmo com a tela inteira dizendo "ambiente de teste" - exatamente o
# tipo de confusão que o modo teste existe pra impedir.
COMANDOS_COM_TESTE = ("sync", "conferir", "concluir")


def modo_teste_pedido(argv: list[str] | None) -> bool:
    """True se este ``argv`` pede o modo teste (sandbox) — regra única.

    Pura: recebe o ``sys.argv`` cru (com o nome do script na posição 0). Só vale
    pros subcomandos que leem/escrevem as planilhas do mês (``sync``/``conferir``);
    nos outros a flag nem existe.
    """
    argumentos = list(argv or [])
    comando = argumentos[1:2]
    if not comando or comando[0] not in COMANDOS_COM_TESTE:
        return False
    return FLAG_TESTE in argumentos[2:]


def modo_teste_ativo(base: Path) -> bool:
    """True se ``base`` (o ``config.PATH_BASE`` em uso) é o sandbox.

    É a checagem do outro lado: não "pediram teste?", e sim "os caminhos que o
    programa VAI usar são mesmo os da cópia?". Serve pro `cli` recusar rodar
    quando as duas pontas discordam.
    """
    try:
        return Path(base).resolve() == PASTA_SANDBOX.resolve()
    except OSError:  # caminho inválido/inacessível: no mínimo não é o sandbox
        return False

# Subpastas que o pipeline espera achar dentro da base (PATH_BASE).
_SUBPASTAS = ("Planos e Macs",)

# Subpastas de dados que também precisam vir pro sandbox mas NÃO ficam no topo
# de `_SUBPASTAS` (caminho relativo à raiz do projeto). Sem isto o `conferir
# --teste` nasce sem os POPs pra conferir (B-QA-2). Podem não existir - aí a
# gente só ignora.
_SUBPASTAS_EXTRA = ("Planos e Macs/POPs do Sistema",)


def _sandbox_tem_planilhas() -> bool:
    """True se já existe pelo menos um .xlsx em alguma subpasta do sandbox."""
    for sub in _SUBPASTAS:
        pasta = PASTA_SANDBOX / sub
        if pasta.is_dir() and any(pasta.glob("*.xlsx")):
            return True
    return False


def _copiar_planilhas_reais() -> list[str]:
    """Copia os .xlsx reais pras subpastas do sandbox. Devolve o que foi copiado."""
    copiados: list[str] = []
    for sub in (*_SUBPASTAS, *_SUBPASTAS_EXTRA):
        origem = RAIZ_PROJETO / sub
        if not origem.is_dir():
            # As de topo têm que existir; as extras (ex.: POPs do Sistema) não.
            if sub in _SUBPASTAS:
                log.warning("Pasta real '%s' não encontrada em %s.", sub, RAIZ_PROJETO)
            continue
        destino = PASTA_SANDBOX / sub
        destino.mkdir(parents=True, exist_ok=True)
        for arquivo in sorted(origem.glob("*.xlsx")):
            if arquivo.name.startswith("~$"):
                continue  # arquivo de lock temporário do Excel
            shutil.copy2(arquivo, destino / arquivo.name)
            copiados.append(f"{sub}/{arquivo.name}")
    return copiados


def preparar_sandbox(forcar_copia: bool = False) -> Path:
    """Garante o sandbox pronto e devolve o caminho dele.

    Só copia as planilhas reais quando o sandbox ainda não tem nenhuma — assim
    dá pra rodar ``sync --teste`` várias vezes seguidas em cima do mesmo estado.
    Pra começar do zero, use ``resetar_sandbox`` (que chama isto com
    ``forcar_copia=True`` pra recopiar mesmo se sobrou coisa que não deu pra
    apagar).
    """
    for sub in _SUBPASTAS:
        (PASTA_SANDBOX / sub).mkdir(parents=True, exist_ok=True)

    if not forcar_copia and _sandbox_tem_planilhas():
        log.info("Ambiente de teste: usando o sandbox existente em %s", PASTA_SANDBOX)
        return PASTA_SANDBOX

    copiados = _copiar_planilhas_reais()
    log.info("Ambiente de teste: sandbox criado em %s", PASTA_SANDBOX)
    for nome in copiados:
        log.info("Ambiente de teste:   copiado %s", nome)
    if not copiados:
        log.warning(
            "Ambiente de teste: nenhuma planilha .xlsx encontrada pra copiar em %s. "
            "O sync vai reclamar de arquivo não encontrado.",
            RAIZ_PROJETO,
        )
    return PASTA_SANDBOX


def _esvaziar_pasta_com_retry(pasta: Path, tentativas: int = 5, espera: float = 0.3) -> list[str]:
    """Apaga tudo DENTRO de ``pasta`` (mantém a raiz). Devolve o que não deu.

    ``shutil.rmtree`` cru estoura ``PermissionError [WinError 5]`` nesta pasta
    OneDrive: o OneDrive/antivírus segura o handle de uma subpasta recém-esvaziada
    por uma fração de segundo e o traceback vai parar na cara do usuário (B-QA-3).
    Aqui a gente vai de baixo pra cima, força permissão de escrita e tenta de
    novo algumas vezes com uma pausa curta. O que sobrar volta na lista (o
    ``resetar_sandbox`` recopia por cima).
    """
    restantes: list[str] = []
    for tentativa in range(1, tentativas + 1):
        restantes = []
        # Do mais fundo pro mais raso: só dá pra remover uma pasta depois de vazia.
        for item in sorted(pasta.rglob("*"), key=lambda p: len(p.parts), reverse=True):
            try:
                if item.is_dir() and not item.is_symlink():
                    item.rmdir()
                else:
                    try:
                        item.chmod(stat.S_IWRITE)
                    except OSError:
                        pass
                    item.unlink()
            except OSError:
                restantes.append(str(item))
        if not restantes:
            return []
        if tentativa < tentativas:
            time.sleep(espera)
    return restantes


def resetar_sandbox() -> Path:
    """Zera o sandbox e recria as cópias do zero.

    Esvazia o conteúdo (sem remover a pasta raiz — ver ``_esvaziar_pasta_com_retry``)
    e recopia as planilhas reais por cima.
    """
    if PASTA_SANDBOX.exists():
        sobrou = _esvaziar_pasta_com_retry(PASTA_SANDBOX)
        if sobrou:
            amostra = ", ".join(sobrou[:5]) + ("..." if len(sobrou) > 5 else "")
            log.warning(
                "Ambiente de teste: %d item(ns) não puderam ser apagados agora "
                "(OneDrive/antivírus segurando?); vou recriar as cópias por cima. %s",
                len(sobrou), amostra,
            )
        else:
            log.info("Ambiente de teste: %s esvaziado.", PASTA_SANDBOX)
    return preparar_sandbox(forcar_copia=True)
