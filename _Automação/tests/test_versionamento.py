"""Testes do prazo por treinamento (R12) — nenhum abre Excel.

A parte que fala com o Excel (`ler_datas_aprovacao`) fica pro teste de
integração; tudo que é REGRA (achar a coluna pelo cabeçalho, montar o
código -> data, somar os 30 dias, apontar quem ficou sem data) é função pura e
está coberto aqui.

Os cabeçalhos usados nos testes são os REAIS, copiados dos arquivos do sandbox
em 2026-09-18 — inclusive o de Junho/2026, que tem duas colunas "Version" e joga
a "Approved Date" pra coluna 8.
"""

from datetime import datetime

import pytest

from treinamentos_its.excel_utils import localizar_coluna
from treinamentos_its.versionamento import (
    CABECALHOS_CODIGO,
    CABECALHOS_DATA_APROVACAO,
    DIAS_DE_PRAZO,
    FORMATO_DIA_PRIMEIRO,
    FORMATO_MES_PRIMEIRO,
    ErroDeLeituraDoVersionamento,
    detectar_formato_das_datas,
    ler_datas_aprovacao,
    localizar_aba_versionamento,
    montar_datas_aprovacao,
    prazo_do_versionamento,
    prazos_por_codigo,
    texto_de_data_e_ambiguo,
)

# Julho/Agosto/Setembro 2026 (e Abril, que só troca "Version" por "Coluna1").
CABECALHO_EN = [
    "Document Number", "Title", "Status", "Version", "Owning Department",
    "Impacted Departments", "Approved Date", "Approver", "Owner", "Reviewer",
    "Observação", "Número do documento anterior",
]
# Junho/2026: DUAS colunas "Version" empurram tudo uma casa pra direita.
CABECALHO_JUNHO = [
    "Document Number", "Title", "Status", "Version", "Version",
    "Owning Department", "Impacted Departments", "Approved Date", "Approver",
    "Owner", "Reviewer", "Observação", "Número do documento anterior",
]
# Janeiro/Fevereiro 2026: o mesmo export, em português.
CABECALHO_PT = [
    "Número do documento", "Título", "Status", "Versão", "Owning Department",
    "Impacted Departments", "Data de aprovação", "Aprovador", "Proprietário",
    "Revisor", "Observação", "Número do documento anterior",
]


# --- achar as colunas pelo cabeçalho -----------------------------------------


def test_coluna_da_data_de_aprovacao_no_layout_mais_comum():
    assert localizar_coluna(CABECALHO_EN, CABECALHOS_DATA_APROVACAO) == 7
    assert localizar_coluna(CABECALHO_EN, CABECALHOS_CODIGO) == 1


def test_coluna_da_data_de_aprovacao_no_layout_de_junho_e_a_8():
    # O motivo de NUNCA usar índice fixo: em Junho/2026 a data está na coluna H.
    # Foi essa diferença que fez o prazo ser descrito ora como "coluna G", ora
    # como "coluna H".
    assert localizar_coluna(CABECALHO_JUNHO, CABECALHOS_DATA_APROVACAO) == 8


def test_colunas_no_layout_em_portugues():
    assert localizar_coluna(CABECALHO_PT, CABECALHOS_CODIGO) == 1
    assert localizar_coluna(CABECALHO_PT, CABECALHOS_DATA_APROVACAO) == 7


def test_cabecalho_desconhecido_nao_chuta_coluna():
    assert localizar_coluna(["Doc", "Data"], CABECALHOS_DATA_APROVACAO) is None


# --- prazo_do_versionamento (a regra: data de aprovação + 30 dias) -----------


def test_prazo_e_a_data_de_aprovacao_mais_30_dias():
    assert prazo_do_versionamento(datetime(2026, 8, 11)) == datetime(2026, 9, 10)


def test_prazo_atravessa_a_virada_de_mes_e_de_ano():
    assert prazo_do_versionamento(datetime(2026, 12, 20)) == datetime(2027, 1, 19)
    # 30 dias corridos, não "um mês": 31/01 + 30 = 02/03 (2026 não é bissexto).
    assert prazo_do_versionamento(datetime(2026, 1, 31)) == datetime(2026, 3, 2)


def test_prazo_zera_a_hora():
    # O Excel devolve data-com-hora; o prazo é um DIA (é assim que ele é
    # comparado com hoje e gravado como dd/mm/yyyy).
    assert prazo_do_versionamento(datetime(2026, 8, 11, 15, 42)) == datetime(2026, 9, 10)


def test_prazo_aceita_data_em_texto():
    # Maio/2026 vem com a maioria das datas como texto no formato americano.
    assert prazo_do_versionamento("4/30/2026") == datetime(2026, 5, 30)
    assert prazo_do_versionamento("11/08/2026") == datetime(2026, 9, 10)


def test_prazo_sem_data_e_none_e_nao_inventa_nada():
    assert prazo_do_versionamento(None) is None
    assert prazo_do_versionamento("") is None
    assert prazo_do_versionamento("sem data") is None


def test_dias_de_prazo_e_trinta():
    # Se alguém mudar isso, é porque o negócio mudou a regra - não por acidente.
    assert DIAS_DE_PRAZO == 30


# --- montar_datas_aprovacao ---------------------------------------------------


def test_montar_datas_ignora_linha_sem_codigo():
    # A aba vem com centenas de linhas vazias no fim (used_range inflado).
    datas, ilegiveis, ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", datetime(2026, 8, 11)), (None, None), ("", datetime(2026, 1, 1))]
    )
    assert datas == {"DOC-POP-0001": datetime(2026, 8, 11)}
    assert ilegiveis == []
    assert ambiguos == []


def test_montar_datas_tira_espaco_do_codigo():
    datas, _ilegiveis, _ambiguos = montar_datas_aprovacao(
        [("  DOC-POP-0001 ", datetime(2026, 8, 11))]
    )
    assert datas == {"DOC-POP-0001": datetime(2026, 8, 11)}


def test_montar_datas_aponta_quem_esta_sem_data_legivel():
    datas, ilegiveis, _ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", None), ("DOC-POP-0002", "xpto"), ("DOC-POP-0003", datetime(2026, 8, 1))]
    )
    assert datas == {"DOC-POP-0003": datetime(2026, 8, 1)}
    assert ilegiveis == ["DOC-POP-0001", "DOC-POP-0002"]


def test_montar_datas_codigo_repetido_vence_a_data_mais_recente():
    # Acontece em poucos códigos por mês nos arquivos reais. A aprovação mais
    # recente é a da versão vigente, que é a que dispara o treinamento.
    datas, _ilegiveis, _ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", datetime(2025, 1, 10)), ("DOC-POP-0001", datetime(2026, 8, 11))]
    )
    assert datas == {"DOC-POP-0001": datetime(2026, 8, 11)}


def test_montar_datas_repetido_fora_de_ordem_tambem_vence_a_mais_recente():
    datas, _ilegiveis, _ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", datetime(2026, 8, 11)), ("DOC-POP-0001", datetime(2025, 1, 10))]
    )
    assert datas == {"DOC-POP-0001": datetime(2026, 8, 11)}


def test_montar_datas_repetido_com_uma_ocorrencia_vazia_nao_vira_ilegivel():
    datas, ilegiveis, _ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", None), ("DOC-POP-0001", datetime(2026, 8, 11))]
    )
    assert datas == {"DOC-POP-0001": datetime(2026, 8, 11)}
    assert ilegiveis == []


def test_montar_datas_com_lista_vazia():
    assert montar_datas_aprovacao([]) == ({}, [], [])
    assert montar_datas_aprovacao(None) == ({}, [], [])


# --- data em texto ambígua: 5/1/2026 é 5 de janeiro ou 1º de maio? -----------


def test_texto_de_data_ambiguo_so_quando_os_dois_cabem_como_mes():
    assert texto_de_data_e_ambiguo("5/1/2026") is True
    assert texto_de_data_e_ambiguo("12/12/2026") is True
    assert texto_de_data_e_ambiguo("4/30/2026") is False  # 30 não é mês
    assert texto_de_data_e_ambiguo("30/04/2026") is False
    assert texto_de_data_e_ambiguo(datetime(2026, 5, 1)) is False  # data de verdade
    assert texto_de_data_e_ambiguo(None) is False
    assert texto_de_data_e_ambiguo("2026-05-01") is False


def test_detectar_formato_pelos_casos_inequivocos():
    # A aba de Maio/2026 é assim: quase tudo em texto no formato americano.
    assert detectar_formato_das_datas(["4/30/2026", "5/1/2026"]) == FORMATO_MES_PRIMEIRO
    assert detectar_formato_das_datas(["30/04/2026", "1/5/2026"]) == FORMATO_DIA_PRIMEIRO


def test_detectar_formato_sem_evidencia_ou_com_evidencia_contraditoria():
    assert detectar_formato_das_datas(["5/1/2026", "12/12/2026"]) is None
    # Planilha misturada: ninguém adivinha (e é melhor recusar do que chutar).
    assert detectar_formato_das_datas(["4/30/2026", "30/04/2026"]) is None
    assert detectar_formato_das_datas([]) is None
    assert detectar_formato_das_datas([datetime(2026, 5, 1), None]) is None


def test_data_ambigua_usa_a_convencao_da_coluna_quando_ela_e_clara():
    # "5/1/2026" numa coluna comprovadamente mm/dd é 1º de MAIO (não 5 de janeiro).
    datas, _ilegiveis, ambiguos = montar_datas_aprovacao(
        [("DOC-POP-0001", "4/30/2026"), ("DOC-POP-0002", "5/1/2026")]
    )
    assert datas["DOC-POP-0002"] == datetime(2026, 5, 1)
    assert ambiguos == []


def test_data_ambigua_sem_convencao_e_recusada_em_vez_de_chutada():
    # Erro aqui desloca um prazo de compliance em MESES: melhor ficar sem prazo
    # e gritar do que gravar uma data errada com cara de certa.
    datas, ilegiveis, ambiguos = montar_datas_aprovacao([("DOC-POP-0001", "5/1/2026")])
    assert datas == {}
    assert ilegiveis == []
    assert ambiguos == ["DOC-POP-0001"]


# --- "não consegui ler" != "não havia nada" ----------------------------------
#
# Workbook dublado (não abre Excel). O que está sob teste é a diferença entre
# devolver {} e ESTOURAR: com {} o `sync` descartava todos os versionamentos "por
# falta de prazo", salvava e dizia que tinha dado certo - sucesso falso.


class _RangeFalso:
    def __init__(self, valor):
        self.value = valor


class _AbaFalsa:
    """Aba de mentira: uma matriz de valores, linha 1 = cabeçalho."""

    def __init__(self, nome: str, linhas: list[list]):
        self.name = nome
        self._linhas = linhas

    @property
    def used_range(self):
        total_linhas = len(self._linhas)
        total_colunas = max((len(linha) for linha in self._linhas), default=0)
        return type(
            "_UsedRange",
            (),
            {"last_cell": type("_Cell", (), {"row": total_linhas, "column": total_colunas})()},
        )()

    def _celula(self, linha: int, coluna: int):
        try:
            return self._linhas[linha - 1][coluna - 1]
        except IndexError:
            return None

    def range(self, inicio, fim):
        (linha1, coluna1), (linha2, coluna2) = inicio, fim
        if linha1 == linha2:
            return _RangeFalso([self._celula(linha1, c) for c in range(coluna1, coluna2 + 1)])
        return _RangeFalso([self._celula(l, coluna1) for l in range(linha1, linha2 + 1)])


class _LivroFalso:
    def __init__(self, abas: list[_AbaFalsa]):
        self.sheets = abas


CABECALHO_COMPLETO = CABECALHO_EN


def _livro_com_versionamento(nome_aba: str, linhas_de_dados: list[list]) -> _LivroFalso:
    return _LivroFalso([_AbaFalsa(nome_aba, [CABECALHO_COMPLETO, *linhas_de_dados])])


def _linha(codigo, data):
    linha = [None] * len(CABECALHO_COMPLETO)
    linha[0] = codigo
    linha[6] = data  # "Approved Date"
    return linha


def test_ler_datas_aprovacao_le_a_aba_normal():
    livro = _livro_com_versionamento(
        "Versionamento Mês", [_linha("DOC-POP-0001", datetime(2026, 8, 11))]
    )
    assert ler_datas_aprovacao(livro, "fake.xlsx") == {"DOC-POP-0001": datetime(2026, 8, 11)}


def test_ler_datas_aprovacao_acha_a_aba_mesmo_com_caixa_e_acento_diferentes():
    livro = _livro_com_versionamento(
        " VERSIONAMENTO MES ", [_linha("DOC-POP-0001", datetime(2026, 8, 11))]
    )
    assert localizar_aba_versionamento(livro) is not None
    assert ler_datas_aprovacao(livro, "fake.xlsx") == {"DOC-POP-0001": datetime(2026, 8, 11)}


def test_ler_datas_aprovacao_sem_a_aba_estoura_em_vez_de_devolver_vazio():
    livro = _LivroFalso([_AbaFalsa("Matriz - Atualização", [["nada"]])])

    with pytest.raises(ErroDeLeituraDoVersionamento) as erro:
        ler_datas_aprovacao(livro, "fake.xlsx")

    assert "Versionamento Mês" in str(erro.value)


def test_ler_datas_aprovacao_com_cabecalho_desconhecido_estoura():
    aba = _AbaFalsa("Versionamento Mês", [["Codigo", "Data"], ["DOC-POP-0001", "01/08/2026"]])

    with pytest.raises(ErroDeLeituraDoVersionamento) as erro:
        ler_datas_aprovacao(_LivroFalso([aba]), "fake.xlsx")

    assert "layout" in str(erro.value).lower()


def test_ler_datas_aprovacao_aba_so_com_cabecalho_nao_estoura():
    # Aba LIDA e sem dados é outra coisa: devolve vazio (com aviso no log) e o
    # ciclo segue - quem não tem data vira "sem prazo" com aviso próprio.
    livro = _livro_com_versionamento("Versionamento Mês", [])
    assert ler_datas_aprovacao(livro, "fake.xlsx") == {}


# --- prazos_por_codigo --------------------------------------------------------


def test_prazos_por_codigo_calcula_um_prazo_por_treinamento():
    datas = {
        "DOC-POP-0000576": datetime(2026, 7, 3),
        "DOC-POP-0028346": datetime(2026, 8, 24),
    }
    prazos, sem_prazo = prazos_por_codigo(["DOC-POP-0000576", "DOC-POP-0028346"], datas)

    # Cenário real do ciclo Julho->Agosto no sandbox: prazos DIFERENTES pro mesmo
    # ciclo - é exatamente o que a regra antiga (uma data só) não conseguia.
    assert prazos == {
        "DOC-POP-0000576": datetime(2026, 8, 2),
        "DOC-POP-0028346": datetime(2026, 9, 23),
    }
    assert sem_prazo == []


def test_prazos_por_codigo_lista_quem_nao_esta_na_aba():
    prazos, sem_prazo = prazos_por_codigo(
        ["DOC-POP-0001", "DOC-GOP-0000802"], {"DOC-POP-0001": datetime(2026, 8, 11)}
    )
    assert prazos == {"DOC-POP-0001": datetime(2026, 9, 10)}
    # Quem não tem data fica FORA do dicionário de propósito: quem chama tem que
    # tratar o caso, não receber um prazo chutado.
    assert sem_prazo == ["DOC-GOP-0000802"]


def test_prazos_por_codigo_sem_a_aba_deixa_todo_mundo_sem_prazo():
    # `ler_datas_aprovacao` devolve {} quando a aba/coluna não é encontrada.
    prazos, sem_prazo = prazos_por_codigo(["DOC-POP-0001", "DOC-POP-0002"], {})
    assert prazos == {}
    assert sem_prazo == ["DOC-POP-0001", "DOC-POP-0002"]


def test_prazos_por_codigo_nao_repete_codigo_na_lista_de_sem_prazo():
    _prazos, sem_prazo = prazos_por_codigo(["DOC-POP-0001", "DOC-POP-0001"], {})
    assert sem_prazo == ["DOC-POP-0001"]


def test_prazos_por_codigo_ignora_entrada_vazia():
    prazos, sem_prazo = prazos_por_codigo([None, "", "  "], {})
    assert prazos == {}
    assert sem_prazo == []
