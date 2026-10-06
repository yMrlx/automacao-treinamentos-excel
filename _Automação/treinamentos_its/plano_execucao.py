"""O que cada botão da GUI vai rodar, e em quais arquivos vai mexer.

Quatro coisas moram aqui, todas **puras** (sem Excel, sem tkinter, sem disco):

1. **Montagem do comando** — a lista de argumentos que vai depois de ``main.py``.
   A GUI não inventa comando nenhum: ela chama ``argumentos_do_comando`` e passa
   o resultado pro subprocesso. Assim o que a janela executa é exatamente o que o
   ``.bat`` executa, e o ``--teste`` sai escrito por inteiro — nunca abreviado
   (foi uma abreviação aceita pelo argparse, ``--tes``, que fez um "modo teste"
   rodar contra as planilhas de PRODUÇÃO).

2. **O plano de arquivos e a prévia** — quais arquivos a ação vai LER e quais vai
   CRIAR/ALTERAR, e o mesmo em português de gente (``resumo_humano``). O cálculo
   é feito a partir de uma ``base`` recebida por parâmetro, e não do
   ``config.PATH_BASE``: o ``config`` resolve os caminhos no import, então dentro
   do processo da GUI ele sempre aponta pra produção. Passar a base explícita é o
   que permite mostrar o caminho certo do sandbox sem mentir.

3. **O que o painel se recusa a rodar** (``motivo_para_nao_rodar``) — hoje, o
   ciclo e o conferir em PRODUÇÃO sem o export de POPs vigentes.

4. **O andamento e o desfecho** — em que etapa a execução está (pelo texto do
   log) e, no fim, o que dizer pra operadora (``desfecho_da_execucao``). Código
   de saída 0 NÃO é sinônimo de "fiz o trabalho": a guarda de reprocessamento
   respondida "n" também sai com 0.

Por que a GUI roda um **subprocesso** em vez de importar a ``cli``: o
``config.PATH_BASE`` é resolvido uma vez só, no import. Num processo só, alternar
entre teste e produção exigiria reimportar meio pacote — e bastaria um import
esquecido pra janela dizer "teste" enquanto escreve em produção. Cada execução
sai num processo novo, que nasce com a base certa e passa pelas mesmas travas do
``main.py``/``cli`` (``modo_teste_pedido`` e ``_abortar_se_teste_sem_sandbox``).
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path, PurePath
from typing import Iterable

from treinamentos_its.ambiente_teste import PASTA_SANDBOX, RAIZ_PROJETO

MODO_TESTE = "teste"
MODO_PRODUCAO = "producao"

ACAO_CONFERIR = "conferir"
ACAO_SYNC = "sync"
ACAO_CONCLUIR = "concluir"
ACAO_RESETAR_TESTE = "resetar-teste"

# Ações que ESCREVEM em arquivo. Em produção, a GUI exige confirmação explícita
# pra qualquer uma delas.
ACOES_QUE_ESCREVEM = (ACAO_SYNC, ACAO_CONCLUIR, ACAO_RESETAR_TESTE)

# Ações que aceitam a flag --teste. Tem que bater com
# `ambiente_teste.COMANDOS_COM_TESTE`, senão a GUI manda uma flag que o argparse
# recusa (ou, pior, que o main.py ignora na hora de preparar o sandbox).
ACOES_COM_TESTE = (ACAO_SYNC, ACAO_CONFERIR, ACAO_CONCLUIR)

# Ações em que o seletor "mês do ciclo" da janela vale. No `sync` e no
# `conferir`, `--mes/--ano` é o mês dos POPs = o do arquivo CRIADO (o ciclo parte
# do Planos e Macs do mês anterior a ele; regra do negócio, 06/10/2026 - antes era o
# mês de ORIGEM). No `concluir` o mesmo `--mes/--ano` quer dizer outra coisa: é o
# arquivo que vai ser ALTERADO. Mandar o mês do ciclo pro `concluir` limpava a
# pessoa no Planos e Macs do mês anterior - entregável já fechado. Por isso o
# `concluir` nunca recebe o mês do seletor: o backend usa o mês da `--data`.
ACOES_COM_MES = (ACAO_SYNC, ACAO_CONFERIR)

# Decisão de negócio (25/09/2026): em PRODUÇÃO o painel só roda o ciclo e o
# conferir com o export de POPs vigentes escolhido. Sem ele, o ciclo cai no
# formato antigo (Matriz do mês anterior como ela está) - que continua existindo
# pelos `.bat`, mas não é o que o painel deve fazer com as planilhas reais.
ACOES_QUE_EXIGEM_VIGENTES_EM_PRODUCAO = (ACAO_SYNC, ACAO_CONFERIR)

# Mesma ordem e grafia de `config.MESES` (há teste garantindo). Não importo o
# `config` aqui porque ele resolve os caminhos de produção no import, e este
# módulo é justamente o que precisa funcionar sem saber a base.
_NOMES_DOS_MESES = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)


def base_do_modo(modo: str) -> Path:
    """A pasta que o processo vai usar como base de dados naquele modo."""
    return PASTA_SANDBOX if modo == MODO_TESTE else RAIZ_PROJETO


def texto_do_banner(modo: str) -> str:
    """A frase do letreiro do topo da janela. PURA."""
    if modo == MODO_TESTE:
        return "AMBIENTE DE TESTE  •  mexe só nas CÓPIAS do sandbox"
    return "PRODUÇÃO  •  mexe nas PLANILHAS REAIS do time"


def _referencia(mes: int | None, ano: int | None) -> list[str]:
    if mes is None or ano is None:
        return []
    return ["--mes", str(mes), "--ano", str(ano)]


def argumentos_do_comando(
    acao: str,
    modo: str,
    *,
    mes: int | None = None,
    ano: int | None = None,
    vigentes: str | Path | None = None,
    obsoletos: str | Path | None = None,
    nome: str | None = None,
    codigo: str | None = None,
    data: str | None = None,
) -> list[str]:
    """Os argumentos que vão DEPOIS de ``main.py``. PURA.

    ``--teste`` só entra nas ações que têm a flag, e sempre escrito por inteiro.
    ``--mes``/``--ano`` só entram juntos (o argparse recusa um sem o outro) e só
    nas ações de ``ACOES_COM_MES``: no ``concluir`` eles são IGNORADOS de
    propósito - ali o mês é o da ``--data`` (ver o comentário de ``ACOES_COM_MES``).
    """
    if acao == ACAO_RESETAR_TESTE:
        return [ACAO_RESETAR_TESTE]

    argumentos: list[str] = [acao]

    if acao == ACAO_CONCLUIR:
        argumentos += [str(nome or ""), str(codigo or "")]
        if data:
            argumentos += ["--data", str(data)]

    if acao in ACOES_COM_MES:
        argumentos += _referencia(mes, ano)

    if acao == ACAO_SYNC and vigentes:
        argumentos += ["--vigentes", str(vigentes)]
    if acao == ACAO_CONFERIR:
        if vigentes:
            argumentos += ["--pops-vigentes", str(vigentes)]
        if obsoletos:
            argumentos += ["--pops-obsoletos", str(obsoletos)]

    if modo == MODO_TESTE and acao in ACOES_COM_TESTE:
        argumentos.append("--teste")
    return argumentos


def acao_escreve_em_arquivo(acao: str) -> bool:
    """True se a ação altera/cria arquivo. PURA."""
    return acao in ACOES_QUE_ESCREVEM


def motivo_para_nao_rodar(
    acao: str, modo: str, vigentes: str | Path | None = None
) -> str | None:
    """Por que o painel NÃO deve rodar esta ação agora, ou ``None`` se pode. PURA.

    Hoje a única recusa é a decisão de negócio: em PRODUÇÃO, ciclo e conferir só
    com o export de POPs vigentes escolhido. Qualquer modo que não seja
    explicitamente TESTE conta como produção (mesma regra de ``base_do_modo``) -
    na dúvida, o lado seguro é recusar.
    """
    if modo == MODO_TESTE or acao not in ACOES_QUE_EXIGEM_VIGENTES_EM_PRODUCAO:
        return None
    if vigentes and str(vigentes).strip():
        return None
    texto = (
        "Em PRODUÇÃO, escolha o arquivo de POPs vigentes que chegou por e-mail "
        "(quadro 2, botão 'Escolher...'). Sem ele o ciclo usaria a Matriz antiga do "
        "mês anterior."
    )
    if acao == ACAO_CONFERIR:
        texto += (
            " A conferência também precisa dele: sem o arquivo ela olharia essa Matriz "
            "antiga, e não a que o ciclo vai usar."
        )
    return texto


def _nomes(meses: dict) -> tuple[str, str]:
    """``(Planos e Macs do mês anterior, Planos e Macs do mês novo)``."""
    anterior = f"Planos e Macs {meses['anterior']} {meses['ano_ant_relativo']}.xlsx"
    atual = f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    return anterior, atual


def plano_de_arquivos(
    acao: str,
    modo: str,
    meses: dict | None = None,
    *,
    vigentes: str | Path | None = None,
    obsoletos: str | Path | None = None,
) -> tuple[list[str], list[str]]:
    """``(o que vai ser LIDO, o que vai ser CRIADO/ALTERADO)``. PURA.

    Caminhos completos, pra a operadora conseguir bater o olho e ver se é o
    sandbox ou a pasta de verdade. A lista de escrita é a que a GUI usa no aviso
    de confirmação — se ela estiver incompleta, o aviso mente, então mexer aqui
    pede o mesmo cuidado de mexer no pipeline.
    """
    base = base_do_modo(modo)
    planos = base / "Planos e Macs"

    if acao == ACAO_RESETAR_TESTE:
        return (
            [f"as planilhas reais em {RAIZ_PROJETO}"],
            [f"{PASTA_SANDBOX} (o sandbox é APAGADO e recriado do zero)"],
        )

    if acao == ACAO_CONCLUIR:
        arquivo = f"{planos}\\Planos e Macs <mês da data> <ano>.xlsx"
        return [arquivo], [arquivo]

    if meses is None:
        return [], []
    nome_anterior, nome_atual = _nomes(meses)

    if acao == ACAO_CONFERIR:
        le = [str(planos / nome_anterior)]
        if vigentes:
            # Este módulo não olha o disco: diz as duas possibilidades.
            le.append(
                f"{planos / nome_atual}  (só se já existir: aí a cópia temporária "
                f"sai DELE, e não do de {_mes_de(meses, 'anterior')})"
            )
            le.append(str(vigentes))
        if obsoletos:
            le.append(str(obsoletos))
        else:
            le.append(
                f"{planos}\\POPs do Sistema\\POPs Obsoletos ... .xlsx  (o mais "
                f"recente da pasta; se não houver, procuro também na raiz de '{planos}')"
            )
        return le, [
            "(nada persistente — o conferir é somente leitura para os arquivos reais; "
            "a cópia de trabalho, quando existe, é temporária e apagada no fim)"
        ]

    if acao == ACAO_SYNC:
        le = [str(planos / nome_anterior)]
        escreve = [f"{planos}\\_pre-colagem\\... (cópia de restauração, se o destino já existir)"]
        if vigentes:
            le.insert(0, str(vigentes))
        escreve.append(
            f"{planos / nome_atual}  (arquivo único do ciclo: Matriz, abas de cargo, "
            "rodapé e versões novas registradas na Matriz)"
        )
        return le, escreve

    return [], []


def resumo_do_plano(
    acao: str,
    modo: str,
    meses: dict | None = None,
    *,
    vigentes: str | Path | None = None,
    obsoletos: str | Path | None = None,
) -> str:
    """O plano em texto, pronto pra caixa de confirmação da GUI. PURA."""
    le, escreve = plano_de_arquivos(acao, modo, meses, vigentes=vigentes, obsoletos=obsoletos)
    linhas = [texto_do_banner(modo), ""]
    linhas.append("VAI LER:")
    linhas += [f"   • {item}" for item in le] or ["   (nada)"]
    linhas.append("")
    linhas.append("VAI CRIAR / ALTERAR:")
    linhas += [f"   • {item}" for item in escreve] or ["   (nada)"]
    return "\n".join(linhas)


# =============================================================================
# O que a OPERADORA lê: antes (prévia) e durante (andamento)
# =============================================================================
#
# A prévia em caminho absoluto saiu da tela principal: quatro linhas de
# ``C:\Users\...\OneDrive - Empresa\...`` não dizem nada pra quem opera,
# estouram a largura da janela e empurram pra fora justamente o que interessa.
# O que fica na tela é o nome do MÊS e o nome do ARQUIVO; a pasta em uso já
# aparece no alto da janela. Os caminhos completos continuam existindo, em
# ``plano_de_arquivos``, atrás de um "ver detalhes" recolhido - servem pra
# diagnóstico, não pro dia a dia.
#
# O que a prévia NÃO pode virar: sumir. Saber o que vai ser alterado ANTES de
# confirmar é o que evita fechar o mês errado, e em produção isso vale dobrado.


def _nome_do_arquivo(caminho: str | Path | None) -> str:
    return Path(str(caminho)).name if caminho else ""


def _mes_de(meses: dict, qual: str) -> str:
    """``'Agosto/2026'`` a partir do dicionário de meses. PURA."""
    if qual == "anterior":
        return f"{meses['anterior']}/{meses['ano_ant_relativo']}"
    return f"{meses['atual']}/{meses['ano_atual']}"


def _mes_da_data(data: str | None) -> str | None:
    """``'Setembro/2026'`` a partir de ``'01/09/2026'``; ``None`` se não der. PURA."""
    try:
        momento = datetime.strptime(str(data or "").strip(), "%d/%m/%Y")
    except ValueError:
        return None
    return f"{_NOMES_DOS_MESES[momento.month - 1]}/{momento.year}"


def _em_producao(modo: str | None) -> bool:
    """Só diz "produção" quando o modo foi informado e não é TESTE."""
    return modo is not None and modo != MODO_TESTE


FRASE_FORMATO_ANTIGO_SYNC = (
    "Sem o arquivo de vigentes: formato ANTIGO — uso a Matriz do mês anterior como ela "
    "está (o mesmo que o 'Executar Ciclo Mensal.bat' faz)."
)


def resumo_humano(
    acao: str,
    meses: dict | None = None,
    *,
    modo: str | None = None,
    vigentes: str | Path | None = None,
    obsoletos: str | Path | None = None,
    nome: str | None = None,
    codigo: str | None = None,
    data: str | None = None,
) -> list[str]:
    """A prévia em português, uma frase por linha, SEM caminho de arquivo. PURA.

    É o texto que a operadora lê antes de clicar, o que aparece na caixa de
    confirmação e o que a GUI escreve na saída no começo de TODA execução. Fala
    em mês ("Agosto/2026") e em nome de arquivo, nunca em caminho - quem precisa
    do caminho abre o "ver detalhes".

    ``modo`` é opcional: sem ele a prévia não afirma nada sobre a trava de
    produção (quem decide se roda é ``motivo_para_nao_rodar``, não este texto).
    """
    if acao == ACAO_RESETAR_TESTE:
        return [
            "Apago o ambiente de teste e refaço as cópias a partir das planilhas reais.",
            "As planilhas reais são só LIDAS - nenhuma delas é alterada.",
        ]

    if acao == ACAO_CONCLUIR:
        quem = nome or "a pessoa informada"
        curso = codigo or "o curso informado"
        mes_da_data = _mes_da_data(data)
        if mes_da_data:
            arquivo = f"no Planos e Macs de {mes_da_data} (o mês da data {data})"
        else:
            arquivo = "no Planos e Macs do mês da data da conclusão"
        return [
            f"Registro a conclusão do curso {curso} para {quem}.",
            f"Mexo somente {arquivo}: gravo OK e esvazio o Planejado em todas as "
            "ocorrências dessa pessoa nesse curso.",
            "O Planos e Macs é sempre o do mês da data da conclusão - o 'mês do ciclo' "
            "escolhido lá em cima não vale aqui.",
            "Se o arquivo ou o par pessoa+curso não existir, termino com erro e não altero nada.",
        ]

    if meses is None:
        return ["Escolha um mês válido pra eu dizer o que vai acontecer."]

    anterior, atual = _mes_de(meses, "anterior"), _mes_de(meses, "atual")

    if acao == ACAO_CONFERIR:
        if vigentes:
            linhas = [
                f"SÓ LEITURA DOS ARQUIVOS REAIS: crio uma cópia temporária do Planos e "
                f"Macs de {atual} se ele já existir (reprocessamento); senão, do de "
                f"{anterior}. Preparo nela a Matriz de {atual} e mostro o que o ciclo "
                f"faria. A cópia é apagada no fim; não altero nem crio entregável.",
                f"Importo e recalculo o export de vigentes "
                f"'{_nome_do_arquivo(vigentes)}' nessa cópia descartável.",
            ]
        elif _em_producao(modo):
            # Nem roda (motivo_para_nao_rodar): não fala de obsoletos.
            return [
                "Em PRODUÇÃO eu só confiro com o arquivo de POPs vigentes escolhido ali "
                "em cima - sem ele eu olharia a Matriz antiga do mês anterior, e não a que "
                "o ciclo vai usar.",
            ]
        else:
            linhas = [
                "SÓ LEITURA DOS ARQUIVOS REAIS: mostro o que o ciclo faria com a Matriz. "
                "Não altero nem crio entregável.",
                f"Sem o arquivo de vigentes: formato ANTIGO — confiro a Matriz do Planos e "
                f"Macs de {anterior} como ela está (a mesma que o 'Executar Ciclo "
                f"Mensal.bat' usa).",
            ]
        if obsoletos:
            linhas.append(f"E comparo com o de obsoletos '{_nome_do_arquivo(obsoletos)}'.")
        else:
            linhas.append(
                "E comparo com o export de obsoletos mais recente da pasta 'POPs do "
                "Sistema'; se não houver ali, procuro também na raiz de 'Planos e Macs'."
            )
        return linhas

    if acao == ACAO_SYNC:
        linhas = [
            f"Leio o Planos e Macs de {anterior} e crio (ou reaproveito, se já existir) "
            f"o de {atual}."
        ]
        if vigentes:
            linhas.append(
                f"No arquivo NOVO, troco a aba 'Versionamento Mês' pelo export "
                f"'{_nome_do_arquivo(vigentes)}' e confiro se a Matriz recalculou."
            )
        elif _em_producao(modo):
            linhas.append(
                "Em PRODUÇÃO eu só rodo com o arquivo de POPs vigentes escolhido ali em "
                "cima - sem ele o ciclo usaria a Matriz antiga do mês anterior."
            )
        else:
            linhas.append(FRASE_FORMATO_ANTIGO_SYNC)
        linhas.append(
            f"Atualizo diretamente os prazos, os status e os rodapés nas abas de cargo "
            f"do Planos e Macs de {atual}. Não escrevo na linha 'Versão' dos blocos: "
            f"ela é fórmula sobre a Matriz."
        )
        linhas.append(
            "No fim, registro na Matriz a versão nova como a versão atual (a coluna E "
            "recebe a versão da F), só onde F é número de versão. É isso que muda a "
            "versão que os blocos mostram e evita repetir o aviso no mês que vem."
        )
        linhas.append(
            f"Se o Planos e Macs de {atual} já existir, tiro cópia de segurança dele "
            f"antes de mexer e, se algo falhar, ele volta como estava. Se ele nascer "
            f"nesta execução e algo falhar, ele é apagado. O de {anterior} nunca é "
            f"alterado."
        )
        return linhas

    return []


# --- andamento: da linha de log pra "em que pé estamos" ----------------------
#
# Os marcadores abaixo são pedaços de mensagens que o backend JÁ escreve. Se
# alguém mudar o texto de um `log.info` lá dentro, a etapa correspondente para
# de acender. Há teste (`test_plano_execucao`) que lê o código do backend e
# exige cada marcador literalmente lá - renomeou a mensagem, o teste quebra.
#
# O marcador de uma etapa é a mensagem que diz que ela COMEÇOU (ou que a
# anterior acabou). Etapa que nunca acendeu não é pintada como feita no fim: ela
# fica com a marca de "não chegou a rodar" (ver `estados_das_etapas`).
MARCADOR_VIGENTES = "IMPORTANDO POPs VIGENTES"

ETAPAS_SYNC = (
    ("Cópia de restauração de", "Tirando cópia de segurança"),
    (MARCADOR_VIGENTES, "Colando o export de POPs vigentes"),
    ("Analisando a Matriz em:", "Validando a Matriz"),
    ("Matriz avaliada:", "Atualizando as abas de cargo"),
    ("Versões novas registradas na Matriz", "Registrando as versões novas na Matriz"),
    ("Planos e Macs conferido após reabrir", "Conferindo o arquivo salvo"),
)
ETAPAS_CONFERIR = (
    (MARCADOR_VIGENTES, "Preparando a cópia descartável do mês novo"),
    ("Analisando a Matriz em:", "Lendo e validando a Matriz"),
    ("Conferência concluída", "Finalizando a conferência"),
)
ETAPAS_CONCLUIR = (
    ("Cópia de restauração de", "Tirando cópia de segurança"),
    ("Conclusão confirmada", "Conferindo as planilhas depois de salvar"),
)
ETAPAS_RESETAR = (
    ("esvaziado", "Esvaziando o ambiente de teste"),
    ("sandbox criado", "Copiando as planilhas reais"),
)

_TABELA_ETAPAS = {
    ACAO_SYNC: ETAPAS_SYNC,
    ACAO_CONFERIR: ETAPAS_CONFERIR,
    ACAO_CONCLUIR: ETAPAS_CONCLUIR,
    ACAO_RESETAR_TESTE: ETAPAS_RESETAR,
}


def etapas_da_acao(acao: str, com_vigentes: bool = False) -> list[str]:
    """Os rótulos das etapas daquela ação, na ordem. PURA.

    Sem vigentes a etapa da colagem nem aparece - etapa que nunca vai acender só
    faria a operadora achar que travou.
    """
    return [
        rotulo
        for marcador, rotulo in _TABELA_ETAPAS.get(acao, ())
        if com_vigentes or marcador != MARCADOR_VIGENTES
    ]


def indice_da_etapa(acao: str, linha: str, com_vigentes: bool = False) -> int | None:
    """Em que etapa esta linha de log coloca a execução (0-based), ou ``None``. PURA."""
    rotulos = etapas_da_acao(acao, com_vigentes)
    for marcador, rotulo in _TABELA_ETAPAS.get(acao, ()):
        if marcador in (linha or "") and rotulo in rotulos:
            return rotulos.index(rotulo)
    return None


# Estado de cada etapa na tela. A GUI traduz em símbolo e cor.
ESTADO_PENDENTE = "pendente"            # ainda não chegou lá
ESTADO_ATUAL = "atual"                  # é aqui que a execução está
ESTADO_FEITA = "feita"                  # acendeu e a execução passou dela
ESTADO_NAO_EXECUTADA = "nao_executada"  # nunca acendeu e não vai mais acender
ESTADO_ERRO = "erro"                    # onde a execução parou com erro


def estados_das_etapas(
    total: int,
    acesas: Iterable[int],
    *,
    terminou: bool = False,
    sucesso: bool = True,
) -> list[str]:
    """O estado de cada uma das ``total`` etapas. PURA.

    ``acesas`` = índices das etapas cujo marcador apareceu no log. A regra que
    importa é a do fim com sucesso: **só é "feita" a etapa que acendeu**. Antes,
    qualquer exit 0 pintava tudo de ✔ - inclusive quando a operadora respondeu
    "n" na guarda de reprocessamento e nada rodou.

    - rodando: antes da atual = feita (se acendeu) ou não executada; a atual =
      atual; depois = pendente;
    - terminou bem: feita (se acendeu) ou não executada;
    - terminou com erro: a última que acendeu (ou a primeira, se nenhuma
      acendeu) = erro; antes dela, feita/não executada; depois, não executada.
    """
    acesas = {indice for indice in acesas if 0 <= indice < total}
    atual = max(acesas) if acesas else -1
    estados: list[str] = []
    for indice in range(total):
        acendeu = indice in acesas
        if not terminou:
            if indice == atual:
                estados.append(ESTADO_ATUAL)
            elif indice < atual:
                estados.append(ESTADO_FEITA if acendeu else ESTADO_NAO_EXECUTADA)
            else:
                estados.append(ESTADO_PENDENTE)
        elif sucesso:
            estados.append(ESTADO_FEITA if acendeu else ESTADO_NAO_EXECUTADA)
        else:
            ponto_do_erro = max(atual, 0)
            if indice == ponto_do_erro:
                estados.append(ESTADO_ERRO)
            elif indice < ponto_do_erro:
                estados.append(ESTADO_FEITA if acendeu else ESTADO_NAO_EXECUTADA)
            else:
                estados.append(ESTADO_NAO_EXECUTADA)
    return estados


# --- desfecho: o que dizer no fim --------------------------------------------
#
# Exit 0 quer dizer "não deu erro", não "fiz o trabalho". Os três marcadores
# abaixo mudam o que a tela diz no fim; todos são mensagens que o backend já
# escreve (e o mesmo teste dos marcadores de etapa exige cada um no código).
MARCADOR_NADA_ALTERADO = "Nada foi alterado"          # cli: guarda respondida "n"
# cli/conclusao: o desfazer FINAL não voltou todos os arquivos. É a palavra final
# do backend. Um "RESTAURAÇÃO FALHOU" solto NÃO basta: ele pode vir da transação
# de uma etapa e a transação do ciclo inteiro, logo depois, restaurar tudo.
MARCADOR_RESTAURACAO_INCOMPLETA = "RESTAURAÇÃO INCOMPLETA"
# planos_macs: resumo final dos status preservados para revisão manual.
MARCADOR_STATUS_DESCONHECIDOS = "status desconhecido(s)"
# excel_app: quit() e kill() falharam - ficou um Excel invisível vivo, que pode
# segurar a planilha na próxima execução (revisão do Codex, 05/10/2026).
MARCADOR_EXCEL_NAO_ENCERRADO = "NÃO consegui encerrar o Excel invisível"
# planos_macs: o ciclo GRAVOU e conferiu reabrindo o arquivo. É a prova positiva
# de gravação que a oferta de abrir o arquivo exige (revisão do Codex, 06/10/2026).
MARCADOR_PLANOS_GRAVADO = "Planos e Macs atualizado com sucesso"

# Mesmo valor de `backup.CODIGO_SAIDA_RESTAURACAO_INCOMPLETA` (repetido aqui pra
# este módulo continuar puro - importar o backup puxaria o Excel pro painel). Um
# teste garante que os dois não se separam.
CODIGO_SAIDA_RESTAURACAO_INCOMPLETA = 3

MARCADORES_DE_DESFECHO = (
    MARCADOR_NADA_ALTERADO,
    MARCADOR_RESTAURACAO_INCOMPLETA,
    MARCADOR_STATUS_DESCONHECIDOS,
    MARCADOR_EXCEL_NAO_ENCERRADO,
    MARCADOR_PLANOS_GRAVADO,
)

NIVEL_OK = "ok"
NIVEL_NEUTRO = "neutro"
NIVEL_ALERTA = "alerta"
NIVEL_ERRO = "erro"
NIVEL_ERRO_GRAVE = "erro_grave"

# Níveis em que a execução NÃO parou com erro (as etapas acesas viram ✔).
NIVEIS_SEM_ERRO = (NIVEL_OK, NIVEL_NEUTRO, NIVEL_ALERTA)

TEXTO_OK = "Terminou sem erros."
TEXTO_NADA_ALTERADO = "Nada foi alterado."
TEXTO_STATUS_DESCONHECIDOS = (
    "Terminou, mas leia os avisos do fim: há status que o programa não conhece "
    "e não mexeu."
)
TEXTO_EXCEL_NAO_ENCERRADO = (
    "Terminou, mas ficou um Excel invisível aberto: antes de rodar de novo, feche "
    "pelo Gerenciador de Tarefas o EXCEL.EXE que não tem janela."
)
TEXTO_RESTAURACAO_FALHOU = (
    "O programa NÃO conseguiu desfazer tudo — veja as linhas RESTAURAÇÃO "
    "INCOMPLETA / RESTAURAÇÃO FALHOU e restaure à mão pela cópia indicada"
)


def desfecho_da_execucao(
    acao: str, codigo_saida: int, linhas_de_log: Iterable[str]
) -> tuple[str, str]:
    """``(texto pra operadora, nível)`` do fim de uma execução. PURA.

    Em ordem de prioridade:

    1. Código 3 ou ``RESTAURAÇÃO INCOMPLETA`` no log → erro GRAVE: o desfazer
       final não voltou todos os arquivos, algum pode ter ficado pela metade.
       (Um ``RESTAURAÇÃO FALHOU`` solto, com outro código, é de uma etapa que a
       transação do ciclo consertou depois — erro normal, não grave.)
    2. Código ≠ 0 → erro normal.
    3. Excel invisível que não morreu (código 0, mas o log diz que o ``kill``
       falhou) → alerta: o resultado vale, mas a próxima execução pode tropeçar
       nele. Antes caía em "Terminou sem erros" (revisão do Codex, 05/10/2026).
    4. Ação que ESCREVE com ``Nada foi alterado`` → neutro (a guarda de
       reprocessamento respondida "n" sai com 0 sem mexer em nada). O
       ``conferir`` fica de fora de propósito: ele é só leitura e termina TODA
       conferência bem-sucedida com essa mesma frase.
    5. Ciclo com aviso final de status desconhecidos → alerta: terminou, mas
       preservou células que a operadora precisa conferir à mão.
    6. Senão → terminou sem erros.
    """
    linhas = [str(linha or "") for linha in linhas_de_log]

    def tem(marcador: str) -> bool:
        return any(marcador in linha for linha in linhas)

    # O 3 só tem esse significado nas ações com transação (ciclo e concluir);
    # nas outras é um erro qualquer.
    restauracao_incompleta = (
        acao in (ACAO_SYNC, ACAO_CONCLUIR)
        and codigo_saida == CODIGO_SAIDA_RESTAURACAO_INCOMPLETA
    )
    if restauracao_incompleta or tem(MARCADOR_RESTAURACAO_INCOMPLETA):
        return TEXTO_RESTAURACAO_FALHOU, NIVEL_ERRO_GRAVE
    if codigo_saida != 0:
        return f"Terminou com erro (código {codigo_saida}).", NIVEL_ERRO
    if tem(MARCADOR_EXCEL_NAO_ENCERRADO):
        return TEXTO_EXCEL_NAO_ENCERRADO, NIVEL_ALERTA
    if acao_escreve_em_arquivo(acao) and tem(MARCADOR_NADA_ALTERADO):
        return TEXTO_NADA_ALTERADO, NIVEL_NEUTRO
    if acao == ACAO_SYNC and tem(MARCADOR_STATUS_DESCONHECIDOS):
        return TEXTO_STATUS_DESCONHECIDOS, NIVEL_ALERTA
    return TEXTO_OK, NIVEL_OK


# Caracteres proibidos em nome de arquivo no Windows, MAIS aspas, parênteses,
# vírgula e ponto e vírgula: sem excluir esses, a expressão atravessava o
# `X.xlsx (cópia de 'Y.xlsx')` e devolvia os dois nomes grudados.
_PROIBIDOS = r"[^\s\/:*?\"<>|'(),;]+"
_PADRAO_ARQUIVO = re.compile(
    _PROIBIDOS + r"(?:[ ]" + _PROIBIDOS + r")*\.xlsx", re.IGNORECASE
)


def arquivo_da_linha(linha: str) -> str | None:
    """O nome do PRIMEIRO ``.xlsx`` citado na linha, sem o caminho. PURA.

    Serve pro "o que estou mexendo agora" da tela: o log cita caminho completo,
    a operadora só precisa do nome do arquivo. Vale o primeiro porque em
    "Criado: X.xlsx (cópia de 'Y.xlsx')" quem está sendo mexido é o X.
    """
    achado = _PADRAO_ARQUIVO.search(linha or "")
    return PurePath(achado.group(0).strip()).name if achado else None


# --- Pergunta s/n do backend levada pra uma caixa da janela ---

# Formato do logger: "%(asctime)s  %(levelname)-8s  %(message)s". O nível ocupa
# 8 colunas fixas: corta só elas, pra não comer o recuo da mensagem ("  - ...").
_PREFIXO_DO_LOG = re.compile(
    r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}  (?P<nivel>[A-Z][A-Z ]{7})  "
)
_PERGUNTA_S_N = re.compile(r"\(s/n\)\s*:?\s*$")

# Só vale quando o rodapé do mês JÁ está escrito: aí, sem versionamento novo, o
# backend devolve NADA_NOVO e não grava. Se só o arquivo existe (sem o rodapé), o
# ciclo roda inteiro nele (revisão do Codex, 05/10/2026: a caixa prometia demais).
EXPLICACAO_REPROCESSAR = (
    "SIM = segue, mas só grava se aparecer versionamento NOVO; sem nada novo, "
    "termina com \"Nada foi alterado\".\n"
    "NÃO = sai agora, sem alterar nada."
)
EXPLICACAO_REPROCESSAR_SEM_RODAPE = (
    "SIM = roda o ciclo NESTE arquivo que já existe: marca as pessoas, escreve o "
    "rodapé do mês e registra as versões novas na Matriz.\n"
    "NÃO = sai agora, sem alterar nada."
)


def texto_da_pergunta(linhas_de_log: Iterable[str]) -> str:
    """Texto limpo pra caixa de pergunta do painel. PURA.

    O backend pergunta no terminal (``... (s/n): ``) depois de listar os motivos
    em linhas ``WARNING``. A caixa mostrava as últimas 12 linhas da tela inteira
    — prévia, data/hora, nível — e a pergunta se perdia (retorno de uso, 05/10/2026).
    Aqui sai só: os WARNING logo antes da pergunta (sem data/hora/nível), a
    pergunta e, na guarda de reprocessamento, o que SIM e NÃO fazem.
    """
    linhas = [str(linha or "").rstrip() for linha in linhas_de_log]
    indice = next(
        (i for i in range(len(linhas) - 1, -1, -1) if _PERGUNTA_S_N.search(linhas[i])),
        None,
    )
    if indice is None:
        return "\n".join(l for l in linhas[-6:] if l.strip())
    pergunta = _PREFIXO_DO_LOG.sub("", linhas[indice]).strip()
    motivos: list[str] = []
    for linha in reversed(linhas[:indice]):
        if not linha.strip():
            continue
        prefixo = _PREFIXO_DO_LOG.match(linha)
        if not prefixo or prefixo.group("nivel").strip() != "WARNING":
            break
        motivos.append(_PREFIXO_DO_LOG.sub("", linha).rstrip())
    partes = ["\n".join(reversed(motivos))] if motivos else []
    partes.append(pergunta)
    if "Rodar mesmo assim" in pergunta:
        rodape_ja_escrito = any("o rodapé 'Referente" in motivo for motivo in motivos)
        partes.append(
            EXPLICACAO_REPROCESSAR if rodape_ja_escrito else EXPLICACAO_REPROCESSAR_SEM_RODAPE
        )
    return "\n\n".join(partes)


# --- Entradas da operadora que o painel confere antes de rodar (05/10/2026) ---

ANO_MINIMO = 2020
ANO_MAXIMO = 2100
_NOMES_DOS_MESES = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)


def problema_no_mes_escolhido(mes: int | None, ano: int | None) -> str | None:
    """Por que o "mês específico" digitado não vale, ou ``None``. PURA.

    O spinbox do ano vai de 2020 a 2100, mas aceita qualquer coisa digitada: o
    QA de 05/10 passou 9999 e 1900 (o painel deixava e o backend só falhava
    depois, procurando um 'Planos e Macs Agosto 9999.xlsx').
    """
    if mes is None or ano is None:
        return "Preencha o mês (1 a 12) e o ano com números."
    if not 1 <= mes <= 12:
        return "O mês tem que estar entre 1 e 12."
    if not ANO_MINIMO <= ano <= ANO_MAXIMO:
        return f"O ano tem que estar entre {ANO_MINIMO} e {ANO_MAXIMO}."
    return None


def avisos_de_mes_dos_exports(
    meses: dict | None, exports: dict[str, tuple[int | None, int | None]]
) -> list[str]:
    """Exports cujo NOME diz outro mês que não o do ciclo. PURA.

    O export certo é o do MÊS DO CICLO (= do Planos e Macs criado): os POPs de
    Setembro criam o Planos e Macs de Setembro (regra do negócio, 06/10/2026).

    ``exports``: rótulo ("vigentes", "obsoletos") -> ``(mês, ano)`` inferidos do
    nome do arquivo (``vigentes.inferir_mes_ano``). O painel só MOSTRAVA o mês ao
    lado do arquivo; com o export de Setembro no ciclo de Agosto ele rodava sem
    perguntar (QA de 05/10). Nome sem mês identificável não gera aviso aqui (o
    rótulo da tela já diz "mês NÃO identificado").
    """
    if not meses or not exports:
        return []
    ciclo = meses["data_atual"]
    avisos: list[str] = []
    for rotulo, (mes, ano) in exports.items():
        if mes is None:
            continue
        if mes == ciclo.month and (ano is None or ano == ciclo.year):
            continue
        avisos.append(
            f"O export de POPs {rotulo} parece ser de {_NOMES_DOS_MESES[mes - 1]}"
            f"/{ano if ano is not None else '?'}, mas o mês escolhido é "
            f"{meses['atual']}/{meses['ano_atual']} (POPs de {meses['atual']} criam o "
            f"Planos e Macs de {meses['atual']} a partir do de {meses['anterior']})."
        )
    return avisos


# --- Fim do ciclo: oferecer abrir o arquivo criado (pedido da operadora, 06/10/2026) ---


def arquivo_criado_pelo_ciclo(modo: str, meses: dict | None) -> Path | None:
    """O Planos e Macs que o ciclo cria/atualiza naquele modo. PURA."""
    if not meses:
        return None
    return (
        base_do_modo(modo)
        / "Planos e Macs"
        / f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    )


def oferecer_abrir_o_arquivo(
    acao: str, nivel: str, linhas_de_log: Iterable[str]
) -> bool:
    """Depois do ciclo que GRAVOU, perguntar se abre o arquivo. PURA.

    "Terminou sem erros" sozinho não dizia à operadora onde estava o resultado.
    Só no ciclo, sem erro, e só com prova POSITIVA de gravação no log
    (``MARCADOR_PLANOS_GRAVADO``). O nível sozinho não basta: o alerta do Excel
    invisível que não fechou vem antes do "Nada foi alterado" no desfecho, então
    um reprocesso recusado ou desfeito também chegava aqui como alerta
    (revisão do Codex, 06/10/2026).
    """
    if acao != ACAO_SYNC or nivel not in (NIVEL_OK, NIVEL_ALERTA):
        return False
    linhas = [str(linha or "") for linha in linhas_de_log]
    if any(MARCADOR_NADA_ALTERADO in linha for linha in linhas):
        return False
    return any(MARCADOR_PLANOS_GRAVADO in linha for linha in linhas)
