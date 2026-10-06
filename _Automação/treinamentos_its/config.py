"""Configurações centralizadas da automação.

Os caminhos são calculados a partir deste arquivo para que o projeto possa ser
movido sem editar o código. Para processar um mês diferente do padrão, prefira
usar ``sync --mes <1-12> --ano <aaaa>`` no terminal.
"""

import os
from datetime import datetime
from pathlib import Path


# Este arquivo fica em:
#   Automatização_Tabelas_Treinamentos/_Automação/treinamentos_its/config.py
# As planilhas "Planos e Macs" ficam na raiz do projeto, DOIS níveis acima
# desta pasta (fora de "_Automação/"). Por isso parents[2].
#   parents[0] = treinamentos_its
#   parents[1] = _Automação
#   parents[2] = Automatização_Tabelas_Treinamentos  <- raiz do projeto
#
# Exceção: se a env var TREINAMENTOS_ITS_BASE estiver setada e apontar pra uma
# pasta que existe, ela vira a base. É assim que o modo "--teste" da CLI
# redireciona tudo pro sandbox (_Automação/_ambiente-teste/) sem encostar nas
# planilhas reais. O main.py seta essa env var ANTES de importar a cli, porque
# estes caminhos são resolvidos aqui no momento do import.
_BASE_ENV = os.environ.get("TREINAMENTOS_ITS_BASE")
if _BASE_ENV and Path(_BASE_ENV).is_dir():
    PATH_BASE = Path(_BASE_ENV).resolve()
else:
    PATH_BASE = Path(__file__).resolve().parents[2]
PASTA_PLANOS_MACS = PATH_BASE / "Planos e Macs"

# Exports do sistema de documentos ("POPs aprovados e efetivos ...", "POPs
# Obsoletos ...") usados só pelo `conferir`. Desde 2026-09-02 ficam numa
# subpasta; o `conferir` procura aqui primeiro e cai na raiz de "Planos e Macs/"
# como compat por 1-2 ciclos.
PASTA_POPS_SISTEMA = PASTA_PLANOS_MACS / "POPs do Sistema"

# Cópias de segurança do "Planos e Macs" tiradas ANTES de a importação dos POPs
# vigentes apagar a aba "Versionamento Mês" inteira. Mesmo espírito do
# Ponto de restauração manual, uma por execução (timestamp no nome). Ninguém
# limpa isso automaticamente.
PASTA_PRE_COLAGEM = PASTA_PLANOS_MACS / "_pre-colagem"


def caminho_copia_pre_colagem(nome_arquivo: str, momento: datetime | None = None) -> Path:
    """Caminho da cópia do Planos e Macs tirada ANTES de recolar o Versionamento.

    Fica em ``Planos e Macs/_pre-colagem/`` com timestamp no nome (``%Y-%m-%d
    %H-%M-%S`` — sem ``:``, que o Windows não aceita em nome de arquivo).
    """
    momento = momento or datetime.now()
    carimbo = momento.strftime("%Y-%m-%d %H-%M-%S")
    caminho = Path(nome_arquivo)
    return PASTA_PRE_COLAGEM / f"{caminho.stem} — antes de {carimbo}{caminho.suffix}"


# None significa: processar o mês anterior ao mês atual. A referência manual
# deve ser fornecida pela CLI para evitar que uma configuração antiga persista.
MES_REFERENCIA: int | None = None
ANO_REFERENCIA: int | None = None

MESES = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)

# Rótulos que aparecem na coluna de nomes das abas de "Planos e Macs" mas NÃO são
# colaboradores (linha de cabeçalho de cada bloco de treinamento). Linhas com
# esses textos são puladas. Se um rótulo real ficar de fora, a linha de cabeçalho
# passa a ser tratada como se fosse uma pessoa.
# Conferido contra "Planos e Macs Julho 2026.xlsx" em 2026-08-31: os rótulos reais
# são "Analista de Sistemas [Junior/Senior]", "Analista ADM", "Analista de
# Suporte", "Gerente IS" e "Estagiário". As comparações ignoram maiúsculas e
# espaços, mas NÃO acentos — por isso as variações com e sem acento.
CABECALHOS_IGNORAR = {
    "Analista de Sistemas",
    "Analista de Sistemas Junior",
    "Analista de Sistemas Júnior",
    "Analista de Sistemas Senior",
    "Analista de Sistemas Sênior",
    "Analista ADM",
    "Analista de Suporte",
    "Gerente IS",
    "Estagiario",
    "Estagiário",
    "ID",
    # rótulos genéricos antigos — mantidos como rede de segurança (não custa
    # ignorar a mais; ninguém se chama assim):
    "Gerente",
    "Terceiros",
}
CABECALHOS_IGNORAR_NORM = {nome.strip().upper() for nome in CABECALHOS_IGNORAR}

# Abas de cargo que contêm os blocos de treinamento. O Planos e Macs é a
# fonte única; não existe mais mapeamento com abas de outra planilha.
ABAS_DE_CARGO = (
    "Analista",
    "Analista Sr",
    "Administrativo",
    "Gerente",
    "Terceiros",
    "Estagiario",
)

LINHA_INICIO_MATRIZ = 5

# Colunas da aba "Matriz - Atualização".
COL_MATRIZ_CODIGO = 3
COL_MATRIZ_VERSAO_ATUAL = 5
COL_MATRIZ_VERSAO_NOVA = 6

STATUS_ON_TIME = "ON TIME"
STATUS_ATRASADO = "ATRASADO"

# Marcador especial da coluna F ("versão nova") da Matriz. Decisão A3
# (2026-09-02, reverte a decisão nº1 de 01/09): "VERIFICAR" (exato, ignorando
# caixa/espaço) significa "POP obsoletado ou renomeado" — NUNCA é versão nova de
# verdade. NÃO versiona ninguém, NÃO escreve nada na coluna de versão (E) do
# Planos. Só gera aviso no terminal + uma linha própria no rodapé de cada aba,
# pedindo pra um humano conferir (ver `dominio.codigos_verificar`). Qualquer
# OUTRO texto não numérico na coluna F continua sendo ignorado (ver
# `dominio.codigos_com_versao_nova_invalida`).
MARCADOR_VERIFICAR = "VERIFICAR"
