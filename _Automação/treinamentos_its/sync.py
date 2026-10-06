"""Leitura e validação do ciclo mensal no Planos e Macs."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from dateutil.relativedelta import relativedelta

from treinamentos_its.config import (
    ABAS_DE_CARGO,
    ANO_REFERENCIA,
    COL_MATRIZ_CODIGO,
    COL_MATRIZ_VERSAO_ATUAL,
    COL_MATRIZ_VERSAO_NOVA,
    LINHA_INICIO_MATRIZ,
    MESES,
    MES_REFERENCIA,
)
from treinamentos_its.dominio import AvaliacaoMatriz, avaliar_matriz
from treinamentos_its.excel_app import abrir_livro, app_excel
from treinamentos_its.logger import log
from treinamentos_its.planos_macs import travas_das_abas_de_cargo
from treinamentos_its.versionamento import (
    DIAS_DE_PRAZO,
    NOME_ABA_VERSIONAMENTO,
    ErroDeLeituraDoVersionamento,
    ler_datas_aprovacao,
)


NOME_ABA_MATRIZ = "Matriz - Atualização"


def obter_meses(
    mes_referencia: int | None = MES_REFERENCIA,
    ano_referencia: int | None = ANO_REFERENCIA,
) -> dict:
    """Mês do ciclo (``atual``) e o mês de onde ele parte (``anterior``).

    **O mês do ciclo é o dos POPs recebidos e o do Planos e Macs que vai ser
    CRIADO** (regra do negócio, 06/10/2026): com os POPs de Setembro, o Planos e
    Macs de Agosto vira o de Setembro, com rodapé "Referente Setembro". Antes o
    mês informado era a ORIGEM e o ciclo ficava um mês adiantado (com os POPs de
    Setembro criava Outubro). Sem informar, o ciclo é o do mês passado — os POPs
    de um mês chegam no começo do mês seguinte.
    """
    if (mes_referencia is None) != (ano_referencia is None):
        raise ValueError("Informe mês e ano juntos, ou não informe nenhum deles.")
    if mes_referencia is not None and ano_referencia is not None:
        if not 1 <= mes_referencia <= 12:
            raise ValueError("O mês de referência deve estar entre 1 e 12.")
        data_atual = datetime(ano_referencia, mes_referencia, 1)
    else:
        hoje = datetime.now()
        data_atual = datetime(hoje.year, hoje.month, 1) - relativedelta(months=1)
    data_anterior = data_atual - relativedelta(months=1)
    log.info(
        "Mês de referência informado: %s/%d (POPs desse mês; cria o Planos e Macs "
        "de %s/%d a partir do de %s/%d).",
        MESES[data_atual.month - 1], data_atual.year,
        MESES[data_atual.month - 1], data_atual.year,
        MESES[data_anterior.month - 1], data_anterior.year,
    )
    return {
        "atual": MESES[data_atual.month - 1],
        "anterior": MESES[data_anterior.month - 1],
        "ano_atual": data_atual.year,
        "ano_ant_relativo": data_anterior.year,
        "data_atual": data_atual,
        "data_anterior": data_anterior,
    }


def _lista_vertical(valor) -> list:
    if valor is None:
        return []
    if not isinstance(valor, list):
        return [valor]
    if valor and isinstance(valor[0], list):
        return [linha[0] if linha else None for linha in valor]
    return valor


def _linhas_de_rodape(wb) -> list[object]:
    """Lê a coluna A das abas de cargo para detectar replay herdado."""
    nomes = {ws.name for ws in wb.sheets}
    linhas: list[object] = []
    for nome_aba in ABAS_DE_CARGO:
        if nome_aba not in nomes:
            continue
        ws = wb.sheets[nome_aba]
        ultima = ws.used_range.last_cell.row
        if ultima < 1:
            continue
        linhas.extend(_lista_vertical(ws.range((1, 1), (ultima, 1)).value))
    return linhas


def avaliar_arquivo_matriz(arquivo: str | Path) -> AvaliacaoMatriz | None:
    """Lê o workbook em somente leitura e chama a avaliação pura única."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        log.error("Arquivo não encontrado: %s", arquivo.resolve())
        return None

    log.info("Analisando a Matriz em: %s", arquivo.resolve())
    try:
        with app_excel() as app:
            wb = abrir_livro(app, arquivo, read_only=True)
            try:
                nomes = {ws.name for ws in wb.sheets}
                if NOME_ABA_MATRIZ not in nomes:
                    log.error("Aba '%s' não encontrada em %s", NOME_ABA_MATRIZ, arquivo.name)
                    return None
                ws = wb.sheets[NOME_ABA_MATRIZ]
                ultima = ws.used_range.last_cell.row
                linhas: list[tuple[str, object, object]] = []
                formulas: list[object] = []
                formulas_e: list[object] = []
                for linha in range(LINHA_INICIO_MATRIZ, ultima + 1):
                    celula_e = ws.range(linha, COL_MATRIZ_VERSAO_ATUAL)
                    celula_f = ws.range(linha, COL_MATRIZ_VERSAO_NOVA)
                    linhas.append(
                        (
                            ws.range(linha, COL_MATRIZ_CODIGO).value,
                            celula_e.value,
                            celula_f.value,
                        )
                    )
                    formulas.append(getattr(celula_f, "formula", None))
                    formulas_e.append(getattr(celula_e, "formula", None))
                rodapes = _linhas_de_rodape(wb)
                # Mesma função que o `atualizar_planos_e_macs` usa antes de
                # escrever: o que para o ciclo lá tem que parar o `conferir` aqui.
                try:
                    travas_abas = travas_das_abas_de_cargo(wb)
                except Exception as erro:  # noqa: BLE001
                    log.exception(
                        "Erro ao ler as abas de cargo de '%s': %s", arquivo, erro
                    )
                    return None
                try:
                    datas_aprovacao = ler_datas_aprovacao(wb, arquivo.name)
                except ErroDeLeituraDoVersionamento as erro:
                    log.error(
                        "O ciclo vai PARAR por isto: não foi possível ler a aba '%s': %s",
                        NOME_ABA_VERSIONAMENTO,
                        erro,
                    )
                    return None
            finally:
                wb.close()
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao ler a Matriz de '%s': %s", arquivo, erro)
        return None

    resultado = avaliar_matriz(
        linhas,
        formulas_f=formulas,
        formulas_e=formulas_e,
        datas_aprovacao=datas_aprovacao,
        linhas_rodape=rodapes,
        travas_abas_de_cargo=travas_abas,
        linha_inicial=LINHA_INICIO_MATRIZ,
        dias_de_prazo=DIAS_DE_PRAZO,
    )
    for aviso in resultado.avisos:
        log.warning("Matriz: %s", aviso)
    for bloqueio in resultado.bloqueios:
        log.error("O ciclo vai PARAR por isto: %s", bloqueio)
    log.info(
        "Matriz avaliada: %d código(s) válido(s), %d versionamento(s), %d aviso(s), "
        "%d bloqueio(s).",
        len(resultado.versoes_matriz),
        len(resultado.alteracoes),
        len(resultado.avisos),
        len(resultado.bloqueios),
    )
    for codigo in sorted(resultado.alteracoes):
        mudanca = resultado.alteracoes[codigo]
        prazo = resultado.prazos.get(codigo)
        aprovado_em = datas_aprovacao.get(codigo)
        # Log em todo ponto de lookup: o prazo tem que dar pra conferir na mão
        # contra a aba 'Versionamento Mês'.
        if prazo and aprovado_em:
            texto_prazo = (
                f"{prazo.strftime('%d/%m/%Y')} (aprovado em "
                f"{aprovado_em.strftime('%d/%m/%Y')} + {DIAS_DE_PRAZO} dias)"
            )
        else:
            texto_prazo = (
                f"indisponível (sem data de aprovação na aba '{NOME_ABA_VERSIONAMENTO}')"
            )
        log.info(
            "Versionamento %s: v%s -> v%s; prazo %s.",
            codigo,
            mudanca["v_antiga"],
            mudanca["v_nova"],
            texto_prazo,
        )
    return resultado
