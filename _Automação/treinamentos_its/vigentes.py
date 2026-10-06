"""Importação do export de **POPs vigentes** para a aba ``Versionamento Mês``.

Hoje isso é feito à mão: alguém apaga a aba inteira e cola o export que chegou
por e-mail. Este módulo faz o mesmo, no arquivo do **mês NOVO**, com três
cuidados que a mão não tem:

1. **Layout canônico.** A coluna F da ``Matriz - Atualização`` é
   ``=SEERRO(PROCV(C5; 'Versionamento Mês'!$A:$D; 4; 0); "Verificar")``. Repare no
   ``4``: a fórmula pega a **quarta coluna, fixa**. O Python acha as colunas pelo
   nome do cabeçalho (``versionamento.py``), mas a planilha não — se o export do
   mês trouxer a versão fora da quarta coluna, a fórmula devolve outra coisa e
   ninguém percebe. Por isso a colagem **remonta** o export na ordem canônica
   (código em A, versão em D, data de aprovação em G), em vez de despejar as
   colunas como vieram. Colunas do export que não são canônicas vão pro fim, pra
   não perder informação.

2. **A aba é limpa por inteiro antes da colagem** (``clear_contents``, que não
   mexe na formatação). Sem isso, um export menor que o do mês passado deixaria
   sobra no fim e o PROCV acharia versão VELHA em vez de devolver ``"Verificar"``.

3. **O recálculo é forçado e CONFERIDO.** Colar e ler acontecem na mesma
   execução; se a leitura vier antes do recálculo, o ciclo inteiro sai errado em
   silêncio. A sequência é: cálculo automático -> ``CalculateFullRebuild()`` ->
   esperar ``CalculationState == xlDone`` (polling com teto, nunca ``sleep`` fixo)
   -> **recalcular o PROCV aqui no Python e comparar célula a célula** com o que
   o Excel devolveu. Só depois disso o arquivo é salvo; se a conferência falhar,
   o workbook é fechado SEM salvar e o arquivo fica como estava.

4. **Deu errado, desfaz.** Antes de encostar na aba, sai uma cópia de segurança
   em ``Planos e Macs/_pre-colagem/``; qualquer caminho de erro restaura o
   arquivo a partir dela. Isso NÃO é redundância: medido em 2026-09-21, nesta
   máquina, fechar o workbook **sem** ``save()`` não desfez a colagem — o arquivo
   apareceu em disco já alterado (AutoSave do OneDrive gravando por baixo). O
   AutoSave é desligado na abertura, mas quem garante é a restauração.

Em que arquivo se cola: **sempre no do mês novo**, nunca no do mês anterior. O
arquivo do mês anterior é entregável de auditoria já fechado e é o ponto de
partida pra refazer o ciclo — colar nele destruiria o ``Versionamento Mês`` que
justificou os versionamentos daquele mês.
"""

from __future__ import annotations

import re
import time
from pathlib import Path

from treinamentos_its.backup import (
    copiar_para_restauracao,
    desligar_autosave,
    restaurar_arquivo,
)
from treinamentos_its.config import (
    COL_MATRIZ_CODIGO,
    COL_MATRIZ_VERSAO_NOVA,
    LINHA_INICIO_MATRIZ,
    MARCADOR_VERIFICAR,
    MESES,
    caminho_copia_pre_colagem,
)
from treinamentos_its.excel_utils import (
    localizar_coluna,
    normalizar_versao,
    sem_acentos,
    versao_como_texto,
)
from treinamentos_its.logger import log
from treinamentos_its.versionamento import (
    CABECALHOS_CODIGO,
    CABECALHOS_DATA_APROVACAO,
    NOME_ABA_VERSIONAMENTO,
    ErroDeLeituraDoVersionamento,
    ler_datas_aprovacao,
    localizar_aba_versionamento,
)

NOME_ABA_MATRIZ = "Matriz - Atualização"

# --- layout canônico da aba "Versionamento Mês" ------------------------------
#
# A ORDEM aqui não é estética: ela é o contrato com a fórmula da coluna F da
# Matriz (``$A:$D`` + índice 4). Mexer nesta tupla sem mexer na fórmula quebra o
# ciclo inteiro em silêncio. Cada item é
# ``(rótulo padrão, nomes aceitos no export, é obrigatória?)``; os nomes aceitos
# estão em ordem de preferência e a busca ignora caixa/acento/espaço
# (``excel_utils.localizar_coluna``).
COLUNAS_CANONICAS: tuple[tuple[str, tuple[str, ...], bool], ...] = (
    ("Document Number", CABECALHOS_CODIGO, True),
    ("Title", ("Title", "Título"), False),
    ("Status", ("Status",), False),
    ("Version", ("Version", "Versão"), True),
    ("Owning Department", ("Owning Department", "Departamento responsável"), False),
    ("Impacted Departments", ("Impacted Departments", "Departamentos impactados"), False),
    ("Approved Date", CABECALHOS_DATA_APROVACAO, True),
    ("Approver", ("Approver", "Aprovador"), False),
)

# Posições que a fórmula da Matriz enxerga (1-based). Conferidas contra os
# arquivos de Janeiro a Setembro/2026: em todos eles o código está em A e a
# versão em D — a fórmula NÃO está lendo coluna errada hoje.
COLUNA_PROCV_CODIGO = 1
COLUNA_PROCV_VERSAO = 4

# Teto pra espera do recálculo do Excel. É polling de estado (CalculationState),
# não espera fixa: quem termina em 2 segundos sai em 2 segundos.
SEGUNDOS_MAX_RECALCULO = 180
INTERVALO_POLL = 0.25

# Constantes do COM do Excel (evita depender das constantes do pywin32).
_XL_CALCULATION_AUTOMATIC = -4105
_XL_DONE = 0


# =============================================================================
# Parte PURA (sem Excel) — tudo aqui é testável sem abrir planilha
# =============================================================================

_MESES_NORMALIZADOS = {
    sem_acentos(nome).casefold(): numero for numero, nome in enumerate(MESES, 1)
}


def inferir_mes_ano(texto: object) -> tuple[int | None, int | None]:
    """``(mês 1-12, ano)`` inferidos do nome do arquivo. PURA.

    A operadora precisa enxergar, ANTES de rodar, que pegou o anexo do mês errado. O
    nome é a única pista que dá pra ler sem abrir o Excel.

    Reconhece, nesta ordem: nome do mês por extenso (com ou sem acento) +
    eventual ano de 4 dígitos; ``mm-aaaa``; ``aaaa-mm``. O que não der pra
    identificar volta como ``None`` — e ``None`` deve virar "não identifiquei" na
    tela, nunca um palpite.
    """
    nome = sem_acentos(str(texto or "")).casefold()

    for rotulo, numero in _MESES_NORMALIZADOS.items():
        posicao = nome.find(rotulo)
        if posicao < 0:
            continue
        # Ano: o 20xx mais próximo depois do nome do mês; se não houver, o
        # primeiro 20xx do texto inteiro.
        depois = re.search(r"(20\d{2})", nome[posicao + len(rotulo):])
        qualquer = re.search(r"(20\d{2})", nome)
        ano = depois or qualquer
        return numero, int(ano.group(1)) if ano else None

    casado = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[-_./](20\d{2})(?!\d)", nome)
    if casado:
        return int(casado.group(1)), int(casado.group(2))
    casado = re.search(r"(?<!\d)(20\d{2})[-_./](0?[1-9]|1[0-2])(?!\d)", nome)
    if casado:
        return int(casado.group(2)), int(casado.group(1))
    return None, None


def descrever_mes_inferido(mes: int | None, ano: int | None) -> str:
    """Texto curto pro rótulo da GUI. PURA."""
    if mes is None:
        return "mês NÃO identificado no nome do arquivo"
    if ano is None:
        return f"parece ser de {MESES[mes - 1]} (ano não identificado)"
    return f"parece ser de {MESES[mes - 1]}/{ano}"


def mapear_colunas_export(cabecalho: list | None) -> tuple[dict[str, int | None], list[str]]:
    """``({rótulo canônico: coluna de origem 1-based}, obrigatórias faltando)``. PURA.

    Coluna não encontrada vira ``None`` (vai vazia pro destino). Obrigatória não
    encontrada entra na segunda lista e a importação **para** — é o caso do
    arquivo de Abril/2026, em que o cabeçalho da versão tinha virado ``'Coluna1'``:
    adivinhar a posição ali seria exatamente a falha silenciosa que a gente
    combate.
    """
    mapa: dict[str, int | None] = {}
    faltando: list[str] = []
    for rotulo, aceitos, obrigatoria in COLUNAS_CANONICAS:
        indice = localizar_coluna(cabecalho, aceitos)
        mapa[rotulo] = indice
        if indice is None and obrigatoria:
            faltando.append(f"{rotulo} (aceito: {' ou '.join(repr(n) for n in aceitos)})")
    return mapa, faltando


def montar_bloco_canonico(
    cabecalho: list | None, linhas: list[list] | None
) -> tuple[list[list], list[str]]:
    """Remonta o export na ordem canônica. PURA.

    Devolve ``(bloco, avisos)``, onde ``bloco[0]`` é o cabeçalho e o resto são os
    dados. As oito colunas canônicas vêm primeiro (código em A, versão em D, data
    em G) e **as colunas restantes do export vão no fim**, com o cabeçalho
    original — a colagem não perde nada, só reordena o que a fórmula enxerga.

    O cabeçalho de cada coluna canônica preserva o texto do export (assim os
    meses em português continuam em português, e o ``versionamento.py`` segue
    achando tudo pelo nome). Só quando a coluna não existe no export é que entra
    o rótulo padrão.

    Linhas sem código são descartadas: o ``used_range`` do Excel infla com linha
    fantasma e elas só serviriam pra empurrar lixo pra aba nova.
    """
    cabecalho = list(cabecalho or [])
    linhas = [list(linha or []) for linha in (linhas or [])]
    mapa, _faltando = mapear_colunas_export(cabecalho)
    avisos: list[str] = []

    usadas = {indice for indice in mapa.values() if indice}
    extras = [
        indice
        for indice in range(1, len(cabecalho) + 1)
        if indice not in usadas and str(cabecalho[indice - 1] or "").strip()
    ]

    def celula(linha: list, indice: int | None):
        if not indice or indice > len(linha):
            return None
        return linha[indice - 1]

    titulos: list[object] = []
    for rotulo, _aceitos, _obrigatoria in COLUNAS_CANONICAS:
        indice = mapa[rotulo]
        if indice is None:
            titulos.append(rotulo)
            avisos.append(f"coluna '{rotulo}' não existe no export - vai VAZIA pra aba")
        else:
            titulos.append(cabecalho[indice - 1])
    titulos.extend(cabecalho[indice - 1] for indice in extras)

    ordem = [mapa[rotulo] for rotulo, _a, _o in COLUNAS_CANONICAS] + extras
    corpo: list[list] = []
    for linha in linhas:
        if not str(celula(linha, mapa["Document Number"]) or "").strip():
            continue
        corpo.append([celula(linha, indice) for indice in ordem])

    origem_codigo = mapa["Document Number"]
    origem_versao = mapa["Version"]
    if origem_codigo not in (None, COLUNA_PROCV_CODIGO) or origem_versao not in (
        None,
        COLUNA_PROCV_VERSAO,
    ):
        avisos.append(
            f"o export veio com o código na coluna {origem_codigo} e a versão na coluna "
            f"{origem_versao}; foram REPOSICIONADOS para {COLUNA_PROCV_CODIGO} e "
            f"{COLUNA_PROCV_VERSAO}, que é o que a fórmula PROCV da Matriz lê"
        )
    for linha in corpo:
        linha[COLUNA_PROCV_VERSAO - 1] = versao_como_texto(linha[COLUNA_PROCV_VERSAO - 1])
    return [titulos, *corpo], avisos



def bloco_para_colar(bloco: list[list]) -> list[list]:
    """Cópia do bloco com todo TEXTO protegido pra colar no Excel. PURA.

    Atribuir ``'1.0'`` ao ``Range.Value`` pelo COM faz o Excel converter pra
    NÚMERO 1 (e ``'5/1/2026'`` viraria data) — a colagem à mão não converte. Em
    05/10/2026 isso deixou a coluna F da Matriz do Setembro com ``1`` em vez de
    ``1.0`` (tudo vermelho na formatação condicional E≠F) e a quitação gravou
    ``4`` na E. O apóstrofo na frente é o "isto é texto" do Excel: some do valor
    (a célula continua valendo ``'1.0'``) e não aparece na tela.
    """
    return [
        [f"'{valor}" if isinstance(valor, str) and valor != "" else valor for valor in linha]
        for linha in bloco
    ]


def versoes_para_procv(bloco: list[list] | None) -> dict[str, object]:
    """``{código em CAIXA ALTA: versão}`` — imita o PROCV do Excel. PURA.

    Duas escolhas copiam o comportamento do Excel de propósito, pra conferência
    valer alguma coisa:

    - **a primeira ocorrência vence** (o PROCV devolve a primeira linha que casa,
      não a última nem a mais recente — diferente de
      ``versionamento.montar_datas_aprovacao``, que para DATA fica com a mais
      recente);
    - **a comparação ignora maiúsculas**, como o PROCV.

    Recebe o bloco já canônico, COM o cabeçalho na primeira posição.
    """
    versoes: dict[str, object] = {}
    for linha in list(bloco or [])[1:]:
        if len(linha) < COLUNA_PROCV_VERSAO:
            continue
        chave = str(linha[COLUNA_PROCV_CODIGO - 1] or "").strip().upper()
        if not chave or chave in versoes:
            continue
        versoes[chave] = linha[COLUNA_PROCV_VERSAO - 1]
    return versoes


def analisar_duplicatas_do_export(
    bloco: list[list] | None,
) -> tuple[list[str], list[tuple[str, list[str]]]]:
    """Devolve ``(duplicados, divergentes)`` do bloco canônico do export.

    O PROCV usa a primeira ocorrência. Duplicata idêntica merece aviso; versões
    diferentes para o mesmo código tornam o resultado ambíguo e devem bloquear a
    importação. Códigos são comparados sem diferenciar caixa.
    """
    vistos: dict[str, list[str]] = {}
    exibicao: dict[str, str] = {}
    duplicados: list[str] = []
    for linha in list(bloco or [])[1:]:
        if len(linha) < COLUNA_PROCV_VERSAO:
            continue
        codigo_original = str(linha[COLUNA_PROCV_CODIGO - 1] or "").strip()
        codigo = codigo_original.upper()
        if not codigo:
            continue
        versao = normalizar_versao(linha[COLUNA_PROCV_VERSAO - 1])
        exibicao.setdefault(codigo, codigo_original)
        if codigo in vistos and exibicao[codigo] not in duplicados:
            duplicados.append(exibicao[codigo])
        if versao not in vistos.setdefault(codigo, []):
            vistos[codigo].append(versao)

    divergentes = [
        (exibicao[codigo], versoes)
        for codigo, versoes in vistos.items()
        if len(versoes) > 1
    ]
    return duplicados, divergentes


def valor_esperado_da_coluna_f(codigo: object, versoes: dict[str, object]) -> str:
    """O que o ``SEERRO(PROCV(...);"Verificar")`` deve devolver pra ``codigo``. PURA."""
    chave = str(codigo or "").strip().upper()
    if not chave:
        return ""
    if chave in versoes:
        return normalizar_versao(versoes[chave])
    return MARCADOR_VERIFICAR


def divergencias_da_coluna_f(
    linhas_matriz: list[tuple[int, object, object, object]] | None,
    versoes: dict[str, object],
) -> tuple[list[str], list[str]]:
    """Confere a coluna F que o Excel devolveu contra o PROCV refeito aqui. PURA.

    ``linhas_matriz`` = ``[(nº da linha, código, valor de F, fórmula de F)]``.
    Devolve ``(divergências, linhas sem fórmula)``:

    - **divergência** em linha COM fórmula é erro duro: significa que o Excel não
      recalculou (ou recalculou em cima de outra coluna). Quem chama aborta.
    - **linha sem fórmula** é aviso: alguém digitou a versão na mão naquela
      célula, então o PROCV não manda nela. Não é erro, mas não pode passar
      calado.
    """
    divergencias: list[str] = []
    sem_formula: list[str] = []
    for numero, codigo, valor, formula in linhas_matriz or []:
        codigo_str = str(codigo or "").strip()
        if not codigo_str:
            continue
        if not str(formula or "").strip().startswith("="):
            sem_formula.append(f"L{numero} {codigo_str} (F='{normalizar_versao(valor)}')")
            continue
        esperado = valor_esperado_da_coluna_f(codigo_str, versoes)
        obtido = normalizar_versao(valor)
        if obtido != esperado:
            divergencias.append(
                f"L{numero} {codigo_str}: o Excel devolveu '{obtido}' e o export diz "
                f"'{esperado}'"
            )
    return divergencias, sem_formula


# =============================================================================
# Parte que toca no Excel
# =============================================================================


def _linhas_do_export(ws) -> tuple[list, list[list]]:
    """``(cabeçalho, linhas)`` da primeira aba do export, numa leitura só."""
    ultima_linha = ws.used_range.last_cell.row
    ultima_coluna = ws.used_range.last_cell.column
    bruto = ws.range((1, 1), (ultima_linha, ultima_coluna)).value
    if not isinstance(bruto, list):
        bruto = [[bruto]]
    elif bruto and not isinstance(bruto[0], list):
        # used_range de uma linha só (lista de colunas) ou de uma coluna só.
        bruto = [bruto] if ultima_linha == 1 else [[valor] for valor in bruto]
    return list(bruto[0]), [list(linha) for linha in bruto[1:]]


def _esperar_calculo(app, segundos: float = SEGUNDOS_MAX_RECALCULO) -> bool:
    """Espera o Excel terminar de calcular. ``False`` se estourar o teto.

    Polling de ``Application.CalculationState`` (``xlDone == 0``) — não é um
    ``sleep`` chutado: se o recálculo acabar no primeiro ciclo, a função volta no
    primeiro ciclo. Se o teto estourar, quem chama **aborta**; seguir daqui seria
    ler valor velho e versionar as pessoas erradas.
    """
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        try:
            estado = app.api.CalculationState
        except Exception as erro:  # noqa: BLE001
            log.warning("Não consegui ler o CalculationState do Excel (%s).", erro)
            return False
        if estado == _XL_DONE:
            return True
        time.sleep(INTERVALO_POLL)
    return False


def _recalcular(app) -> bool:
    """Força o recálculo completo e espera terminar. ``False`` se não deu."""
    try:
        calculo_anterior = app.api.Calculation
    except Exception:  # noqa: BLE001
        calculo_anterior = None
    try:
        if calculo_anterior != _XL_CALCULATION_AUTOMATIC:
            log.info(
                "O Excel estava em cálculo MANUAL (Calculation=%s); ligando o automático "
                "pra esta importação e devolvendo o modo original antes de salvar.",
                calculo_anterior,
            )
            app.api.Calculation = _XL_CALCULATION_AUTOMATIC
        # CalculateFullRebuild refaz a árvore de dependências inteira: é o que
        # garante a reavaliação da fórmula da Matriz depois de a aba de origem ter
        # sido apagada e recolada.
        app.api.CalculateFullRebuild()
        if not _esperar_calculo(app):
            log.error(
                "O Excel não terminou de calcular em %d segundos (CalculationState != "
                "xlDone). NADA foi salvo - ler a Matriz agora daria valor velho.",
                SEGUNDOS_MAX_RECALCULO,
            )
            return False
        return True
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao forçar o recálculo do Excel: %s", erro)
        return False
    finally:
        if calculo_anterior is not None and calculo_anterior != _XL_CALCULATION_AUTOMATIC:
            try:
                app.api.Calculation = calculo_anterior
            except Exception as erro:  # noqa: BLE001
                log.warning("Não consegui devolver o modo de cálculo original: %s", erro)


def _ler_matriz_para_conferencia(wb, rotulo: str) -> list[tuple[int, object, object, object]]:
    """``[(linha, código, valor de F, fórmula de F)]`` da aba Matriz."""
    nomes = [aba.name for aba in wb.sheets]
    if NOME_ABA_MATRIZ not in nomes:
        raise ErroDeLeituraDoVersionamento(
            f"Aba '{NOME_ABA_MATRIZ}' não encontrada em '{rotulo}' - sem ela não dá pra "
            f"conferir se o PROCV recalculou. Abas: {', '.join(nomes)}"
        )
    ws = wb.sheets[NOME_ABA_MATRIZ]
    ultima = ws.used_range.last_cell.row
    saida: list[tuple[int, object, object, object]] = []
    for numero in range(LINHA_INICIO_MATRIZ, ultima + 1):
        codigo = ws.range(numero, COL_MATRIZ_CODIGO).value
        if not str(codigo or "").strip():
            continue
        celula_f = ws.range(numero, COL_MATRIZ_VERSAO_NOVA)
        saida.append((numero, codigo, celula_f.value, celula_f.formula))
    return saida


def _conferir_recalculo(wb, bloco: list[list], rotulo: str) -> bool:
    """Compara a coluna F do Excel com o PROCV refeito no Python. Loga tudo."""
    versoes = versoes_para_procv(bloco)
    try:
        linhas = _ler_matriz_para_conferencia(wb, rotulo)
    except ErroDeLeituraDoVersionamento as erro:
        log.error("%s", erro)
        return False
    if not linhas:
        log.error(
            "A aba '%s' de '%s' não tem nenhuma linha com código a partir da linha %d - "
            "não há o que conferir, e isso não parece certo. Importação ABORTADA.",
            NOME_ABA_MATRIZ, rotulo, LINHA_INICIO_MATRIZ,
        )
        return False

    divergencias, sem_formula = divergencias_da_coluna_f(linhas, versoes)
    if sem_formula:
        log.error(
            "%d linha(s) da Matriz têm a coluna F preenchida NA MÃO (sem fórmula). "
            "Não dá para provar que esses valores vieram do export; a importação foi "
            "BLOQUEADA antes de versionar alguém: %s",
            len(sem_formula),
            "; ".join(sem_formula[:20]) + (" ..." if len(sem_formula) > 20 else ""),
        )
        return False
    if divergencias:
        log.error(
            "RECÁLCULO NÃO CONFERE: %d linha(s) da Matriz estão com a coluna F diferente "
            "do que o export colado manda. Isso significa que o Excel não recalculou (ou "
            "recalculou sobre outra coluna) - seguir daqui versionaria as pessoas erradas. "
            "Divergências: %s",
            len(divergencias),
            "; ".join(divergencias[:20]) + (" ..." if len(divergencias) > 20 else ""),
        )
        return False
    log.info(
        "Recálculo CONFERIDO: as %d linha(s) da Matriz com fórmula batem, uma a uma, com "
        "o PROCV refeito sobre o export colado.",
        len(linhas) - len(sem_formula),
    )
    return True


def _salvar_copia_pre_colagem(
    arquivo: Path, destino_copia: str | Path | None = None
) -> Path | None:
    """Cópia de segurança do Planos e Macs antes de apagar a aba inteira.

    Usa ``shutil.copy2``, nunca
    ``SaveAs`` (que falha de forma determinística dentro do OneDrive desta
    máquina). Devolve o caminho da cópia — ela não é só arquivo de arquivo morto,
    é o que ``_restaurar`` usa pra desfazer uma importação que deu errado.
    """
    destino = (
        Path(destino_copia)
        if destino_copia is not None
        else caminho_copia_pre_colagem(arquivo.name)
    )
    return copiar_para_restauracao(
        arquivo,
        destino,
        descricao=f"Planos e Macs '{arquivo.name}' antes da importação",
    )


def _restaurar(copia: Path, arquivo: Path) -> None:
    """Devolve o arquivo ao estado anterior à colagem.

    **Por que isto existe:** fechar o workbook sem ``save()`` NÃO garante que o
    arquivo em disco ficou intocado. Nesta máquina, com as planilhas dentro da
    árvore do OneDrive, um workbook modificado e fechado sem salvar **apareceu em
    disco com a colagem aplicada** (medido em 2026-09-21: a aba de destino
    passou de 711 pra 707 linhas depois de uma importação que retornou erro).
    É o AutoSave do OneDrive gravando por baixo. A gente desliga o AutoSave ao
    abrir (``_desligar_autosave``), mas não dá pra depender só disso: a
    restauração a partir da cópia é a garantia de verdade.
    """
    restaurar_arquivo(
        copia,
        arquivo,
        descricao=f"Planos e Macs '{arquivo.name}' após falha na importação",
    )


def _desligar_autosave(wb) -> None:
    """Desliga o AutoSave do OneDrive neste workbook, se ele existir.

    Sem isso o Excel vai gravando sozinho enquanto a gente mexe, e uma
    importação abortada no meio deixaria a aba pela metade em disco. A
    propriedade só existe em versões recentes do Excel e só vale pra arquivo em
    OneDrive/SharePoint — por isso tudo aqui é 'se der'.
    """
    desligar_autosave(
        wb,
        descricao=f"'{getattr(wb, 'name', '?')}' durante a importação dos vigentes",
    )


def _colar_e_conferir(app, arquivo_destino: Path, bloco: list[list]) -> bool:
    """Limpa a aba, cola o bloco, recalcula, confere e salva. Não restaura nada."""
    from treinamentos_its.excel_app import abrir_livro

    wb = abrir_livro(app, arquivo_destino)
    try:
        _desligar_autosave(wb)
        ws = localizar_aba_versionamento(wb)
        if ws is None:
            log.error(
                "Aba '%s' não encontrada em '%s'. Abas: %s.",
                NOME_ABA_VERSIONAMENTO,
                arquivo_destino.name,
                ", ".join(aba.name for aba in wb.sheets),
            )
            return False

        linhas_antes = ws.used_range.last_cell.row
        # clear_contents (e não clear): tira o conteúdo da aba inteira e preserva a
        # formatação. Limpar a aba INTEIRA é o ponto - um export menor que o do mês
        # passado deixaria sobra no fim e o PROCV acharia versão VELHA em vez de
        # devolver "Verificar".
        ws.cells.clear_contents()
        log.info(
            "Aba '%s' limpa (tinha %d linha(s) usadas); colando %d linha(s).",
            ws.name, linhas_antes, len(bloco) - 1,
        )
        # Uma atribuição em bloco: uma chamada COM só, não célula a célula. Texto
        # vai protegido (apóstrofo), senão o Excel converte '1.0' em 1.
        ws.range((1, 1)).value = bloco_para_colar(bloco)

        if not _recalcular(app):
            return False
        if not _conferir_recalculo(wb, bloco, arquivo_destino.name):
            return False

        wb.save()
        log.info("Salvo: %s", arquivo_destino.name)
        return True
    finally:
        wb.close()


def _reconferir_do_disco(app, arquivo_destino: Path, bloco: list[list]) -> bool:
    """Reabre o arquivo salvo e confere de novo — pega valor em cache velho."""
    from treinamentos_its.excel_app import abrir_livro

    wb = abrir_livro(app, arquivo_destino, read_only=True)
    try:
        if not _recalcular(app):
            return False
        if not _conferir_recalculo(wb, bloco, arquivo_destino.name):
            log.error(
                "A conferência passou ANTES de salvar e falhou DEPOIS de reabrir '%s'. A "
                "Matriz desse arquivo NÃO pode ser usada como está.",
                arquivo_destino.name,
            )
            return False
        try:
            datas = ler_datas_aprovacao(wb, arquivo_destino.name)
        except ErroDeLeituraDoVersionamento as erro:
            log.error(
                "A aba foi colada mas NÃO dá pra ler a data de aprovação dela: %s. Sem "
                "isso o ciclo não calcula prazo nenhum.",
                erro,
            )
            return False
        log.info(
            "Conferência final: %d código(s) com data de aprovação legível na aba nova "
            "(é dela que sai o prazo de cada treinamento).",
            len(datas),
        )
        return True
    finally:
        wb.close()


def importar_vigentes(
    caminho_export: str | Path,
    arquivo_destino: str | Path,
    *,
    destino_copia_restauracao: str | Path | None = None,
) -> bool:
    """Limpa a aba ``Versionamento Mês`` de ``arquivo_destino`` e cola o export.

    ``arquivo_destino`` é o **Planos e Macs do mês NOVO** (já criado como cópia do
    mês anterior). Retorna ``True`` só quando colou, recalculou **e** conferiu o
    recálculo duas vezes (antes de salvar e depois de reabrir do disco).

    Em qualquer caminho de erro o arquivo é **restaurado** a partir da cópia de
    segurança tirada no começo — ver ``_restaurar`` pra entender por que fechar
    sem salvar não basta nesta máquina. ``destino_copia_restauracao`` existe para
    operações em uma cópia descartável (o ``conferir``): nesse caso até a cópia
    de restauração fica na pasta temporária e some ao final. O ciclo real omite o
    argumento e continua usando ``Planos e Macs/_pre-colagem/``.
    """
    from treinamentos_its.excel_app import abrir_livro, app_excel

    caminho_export = Path(caminho_export)
    arquivo_destino = Path(arquivo_destino)

    if not caminho_export.exists():
        log.error("Export de POPs vigentes não encontrado: %s", caminho_export)
        return False
    if not arquivo_destino.exists():
        log.error("Planos e Macs de destino não encontrado: %s", arquivo_destino)
        return False

    log.info("%s", "-" * 60)
    log.info("IMPORTANDO POPs VIGENTES")
    log.info("   export lido   : %s", caminho_export)
    log.info("   aba reescrita : '%s' de %s", NOME_ABA_VERSIONAMENTO, arquivo_destino.name)

    copia = _salvar_copia_pre_colagem(arquivo_destino, destino_copia_restauracao)
    if copia is None:
        log.error("Sem cópia de segurança eu não apago a aba. Importação interrompida.")
        return False

    sucesso = False
    try:
        with app_excel() as app:
            wb_export = abrir_livro(app, caminho_export, read_only=True)
            try:
                cabecalho, linhas = _linhas_do_export(wb_export.sheets[0])
            finally:
                wb_export.close()

            _mapa, faltando = mapear_colunas_export(cabecalho)
            if faltando:
                log.error(
                    "O export '%s' não tem coluna(s) obrigatória(s): %s. O cabeçalho lido "
                    "foi: %s. NÃO vou adivinhar posição de coluna - a fórmula da Matriz lê "
                    "a 4ª coluna fixa e um palpite errado aqui versiona todo mundo errado. "
                    "Nada foi alterado.",
                    caminho_export.name,
                    "; ".join(faltando),
                    ", ".join(str(t) for t in cabecalho if str(t or "").strip()),
                )
                return False

            bloco, avisos = montar_bloco_canonico(cabecalho, linhas)
            for aviso in avisos:
                log.warning("Export '%s': %s", caminho_export.name, aviso)
            if len(bloco) < 2:
                log.error(
                    "O export '%s' não tem nenhuma linha com código preenchido. Colar isso "
                    "deixaria a Matriz inteira em '%s'. Nada foi alterado.",
                    caminho_export.name, MARCADOR_VERIFICAR,
                )
                return False
            duplicados, divergentes = analisar_duplicatas_do_export(bloco)
            if divergentes:
                log.error(
                    "O export '%s' repete %d código(s) com VERSÕES DIFERENTES: %s. "
                    "O PROCV escolheria só a primeira ocorrência; a importação foi "
                    "bloqueada para não versionar com um resultado ambíguo.",
                    caminho_export.name,
                    len(divergentes),
                    "; ".join(
                        f"{codigo} ({' x '.join(versoes)})"
                        for codigo, versoes in divergentes[:20]
                    ) + (" ..." if len(divergentes) > 20 else ""),
                )
                return False
            if duplicados:
                log.warning(
                    "O export '%s' repete %d código(s), mas com a mesma versão; o PROCV "
                    "usará a primeira ocorrência: %s.",
                    caminho_export.name,
                    len(duplicados),
                    ", ".join(duplicados[:20]) + (" ..." if len(duplicados) > 20 else ""),
                )
            log.info(
                "Export lido: %d linha(s) de dados, %d coluna(s) (%d canônicas + %d "
                "extra(s) preservada(s) no fim).",
                len(bloco) - 1, len(bloco[0]), len(COLUNAS_CANONICAS),
                len(bloco[0]) - len(COLUNAS_CANONICAS),
            )

            sucesso = _colar_e_conferir(app, arquivo_destino, bloco) and _reconferir_do_disco(
                app, arquivo_destino, bloco
            )
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao importar os POPs vigentes: %s", erro)
        sucesso = False

    if not sucesso:
        # A restauração acontece com o Excel JÁ fechado: com o arquivo ainda
        # aberto, o copy2 esbarraria no lock.
        _restaurar(copia, arquivo_destino)
    return sucesso
