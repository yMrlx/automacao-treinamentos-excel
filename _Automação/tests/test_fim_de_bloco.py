"""Fim de bloco e variantes de ELABORADOR/APROVADORA (02/10/2026).

Bug 1: o fim de cada bloco era o ``linha_cod`` do próximo bloco COM código.
As abas reais têm vários blocos "Treinamentos" sem código no formato X-Y-Z
(BPF, Global - Laces, Ilearn...), e as pessoas deles caíam na faixa do bloco
de cima: o ciclo marcava ON TIME em célula de outro curso e o ``concluir``
gravava OK em cursos que a pessoa não fez. Agora todo cabeçalho "Treinamentos"
encerra o bloco de cima (``planos_macs._ler_blocos`` -> ``linha_fim``).

Bug 2: a regra "OK com ELABORADOR/APROVADORA no Planejado não mexe" só
reconhecia o texto exato; agora compara normalizado (strip, caixa, acento) e
aceita as 4 palavras.

Nada aqui abre Excel nem resolve o ``PATH_BASE`` real: o ``config`` aponta pro
``tmp_path`` e ``app_excel``/``abrir_livro`` são dublês (o padrão é EXPLODIR se
algum teste esquecer de trocar).
"""

from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace

import pytest

from treinamentos_its import conclusao, config, planos_macs
from treinamentos_its.dominio import (
    MARCACAO_ON_TIME,
    MARCACAO_PRESERVAR,
    AvaliacaoMatriz,
    decidir_marcacao_planos,
    planejado_de_elaborador_ou_aprovador,
)


HOJE = datetime(2026, 9, 25)
PRAZO = datetime(2026, 10, 20)
MESES = {
    "anterior": "Agosto",
    "ano_ant_relativo": 2026,
    "atual": "Setembro",
    "ano_atual": 2026,
}


# ---------------------------------------------------------------------------
# Dublês de xlwings (mesmo estilo de tests/test_planos_macs.py)
# ---------------------------------------------------------------------------
class _Cell:
    def __init__(self, value=None, formula=None):
        self.value = value
        self.formula = formula
        self.color = None
        self.font = SimpleNamespace(color=None, bold=False, size=None)


class _Range:
    def __init__(self, ws, inicio, fim):
        self.ws = ws
        self.inicio = inicio
        self.fim = fim

    @property
    def value(self):
        li, ci = self.inicio
        lf, cf = self.fim
        dados = [
            [self.ws.range(linha, coluna).value for coluna in range(ci, cf + 1)]
            for linha in range(li, lf + 1)
        ]
        if li == lf:
            return dados[0]
        if ci == cf:
            return [linha[0] for linha in dados]
        return dados


class _Sheet:
    def __init__(self, name, rows, cols):
        self.name = name
        self.cells = {}
        self.used_range = SimpleNamespace(
            last_cell=SimpleNamespace(row=rows, column=cols)
        )
        self.api = SimpleNamespace(AutoFilterMode=False)

    def range(self, inicio, coluna=None):
        if isinstance(inicio, tuple):
            return _Range(self, inicio, coluna)
        return self.cells.setdefault((inicio, coluna), _Cell())

    def valor(self, linha, coluna):
        celula = self.cells.get((linha, coluna))
        return None if celula is None else celula.value


class _Sheets(list):
    def __getitem__(self, key):
        if isinstance(key, str):
            return next(item for item in self if item.name == key)
        return super().__getitem__(key)


class _Workbook:
    def __init__(self, sheets):
        self.sheets = _Sheets(sheets)
        self.api = SimpleNamespace(AutoSaveOn=False)
        self.salvou = False

    def save(self):
        self.salvou = True

    def close(self):
        pass


class _App:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


def _excel_proibido(*_args, **_kwargs):
    raise AssertionError("teste tentou abrir o Excel de verdade")


@pytest.fixture(autouse=True)
def _sem_caminho_real_nem_excel(monkeypatch, tmp_path):
    """Config apontando pro tmp_path e Excel bloqueado até o teste trocar."""
    base = tmp_path / "base"
    pasta_planos = base / "Planos e Macs"
    pasta_planos.mkdir(parents=True)
    monkeypatch.setattr(config, "PATH_BASE", base)
    monkeypatch.setattr(config, "PASTA_PLANOS_MACS", pasta_planos)
    monkeypatch.setattr(config, "PASTA_POPS_SISTEMA", pasta_planos / "POPs do Sistema")
    monkeypatch.setattr(config, "PASTA_PRE_COLAGEM", pasta_planos / "_pre-colagem")
    monkeypatch.setattr(planos_macs, "PASTA_PLANOS_MACS", pasta_planos)
    monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", pasta_planos)
    for modulo in (planos_macs, conclusao):
        monkeypatch.setattr(modulo, "app_excel", _excel_proibido)
        monkeypatch.setattr(modulo, "abrir_livro", _excel_proibido)
    return pasta_planos


def _usar_workbook(monkeypatch, wb, abas):
    for modulo in (planos_macs, conclusao):
        monkeypatch.setattr(modulo, "app_excel", lambda: _App())
        monkeypatch.setattr(modulo, "abrir_livro", lambda *_a, **_k: wb)
        monkeypatch.setattr(modulo, "ABAS_DE_CARGO", tuple(abas))


# ---------------------------------------------------------------------------
# Montagem das abas
# ---------------------------------------------------------------------------
def _bloco(ws, linha_cabecalho, rotulo_curso, pessoas):
    """Escreve um bloco no layout real: cabeçalho, códigos, versão, ID, pessoas.

    ``pessoas``: lista de ``(nome, planejado, status)`` a partir de
    ``linha_cabecalho + 4``; a coluna C é o Planejado e a D o Executado.
    """
    ws.range(linha_cabecalho, 1).value = "Treinamentos"
    ws.range(linha_cabecalho + 1, 1).value = "Códigos dos Treinamentos"
    ws.range(linha_cabecalho + 1, 3).value = rotulo_curso
    ws.range(linha_cabecalho + 2, 1).value = "Versão"
    ws.range(linha_cabecalho + 2, 3).value = "1.0"
    ws.range(linha_cabecalho + 3, 1).value = "ID"
    ws.range(linha_cabecalho + 3, 3).value = "Planejado"
    ws.range(linha_cabecalho + 3, 4).value = "Executado"
    for deslocamento, (nome, planejado, status) in enumerate(pessoas):
        linha = linha_cabecalho + 4 + deslocamento
        ws.range(linha, 2).value = nome
        ws.range(linha, 3).value = planejado
        ws.range(linha, 4).value = status


def _aba_reconhecido_sem_codigo_reconhecido(nome="Cargo", pessoas_sem_codigo=None):
    """(a) bloco DOC-POP-1 (L1) -> bloco SEM código (L8) -> bloco DOC-POP-2 (L14).

    Linhas de pessoa: Ana L5, Bruno L6 | Carla L12 (+ quem vier) | Davi L18.
    """
    ws = _Sheet(nome, rows=18, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Ana", None, "OK"), ("Bruno", None, None)])
    _bloco(
        ws,
        8,
        "Garantia da Qualidade",
        pessoas_sem_codigo or [("Carla", None, "OK")],
    )
    _bloco(ws, 14, "DOC-POP-2", [("Davi", None, "OK")])
    return ws


def _aba_reconhecido_e_sem_codigo_no_fim(nome="Cargo B"):
    """(b) bloco DOC-POP-1 (L1) -> bloco SEM código no fim da aba (L8).

    Linhas de pessoa: Ana L5 | Carla L12 (com OK legítimo do outro curso) e
    Edu L13 (célula vazia, fora de curso).
    """
    ws = _Sheet(nome, rows=13, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Ana", None, "OK")])
    _bloco(
        ws,
        8,
        "Global - Laces",
        [("Carla", None, "OK"), ("Edu", None, None)],
    )
    return ws


def _avaliacao(*codigos):
    codigos = codigos or ("DOC-POP-1",)
    return AvaliacaoMatriz(
        alteracoes={c: {"v_antiga": "1.0", "v_nova": "2.0"} for c in codigos},
        prazos={c: PRAZO for c in codigos},
        linhas_quitacao={c: 5 + i for i, c in enumerate(codigos)},
    )


def _matriz(*codigos):
    codigos = codigos or ("DOC-POP-1",)
    ws = _Sheet("Matriz - Atualização", rows=4 + len(codigos), cols=6)
    for i, codigo in enumerate(codigos):
        linha = 5 + i
        ws.range(linha, 3).value = codigo
        ws.range(linha, 5).value = "1.0"
        ws.range(linha, 6).value = 2.0
        ws.range(linha, 6).formula = "=SEERRO(PROCV())"
    return ws


def _linhas_marcadas(marcacoes):
    return sorted(m.linha for m in marcacoes)


# ---------------------------------------------------------------------------
# _ler_blocos: o fim de cada bloco
# ---------------------------------------------------------------------------
def test_fim_do_bloco_e_o_proximo_cabecalho_mesmo_sem_codigo():
    ws = _aba_reconhecido_sem_codigo_reconhecido()
    blocos, tem_conteudo = planos_macs._ler_blocos(ws)
    assert tem_conteudo
    # O bloco sem código continua NÃO sendo bloco (regra de quais têm código
    # não mudou) — só serve de fronteira.
    assert [b["codigos"] for b in blocos] == [{3: "DOC-POP-1"}, {3: "DOC-POP-2"}]
    assert [(b["linha_cod"], b["linha_func"], b["linha_fim"]) for b in blocos] == [
        (2, 5, 8),  # termina no cabeçalho sem código (L8), não no DOC-POP-2
        (15, 18, 19),  # último: ultima_linha + 1
    ]


def test_ultimo_bloco_termina_no_cabecalho_sem_codigo_do_fim_da_aba():
    ws = _aba_reconhecido_e_sem_codigo_no_fim()
    blocos, _ = planos_macs._ler_blocos(ws)
    assert [(b["linha_cod"], b["linha_fim"]) for b in blocos] == [(2, 8)]


def test_cabecalho_na_ultima_linha_da_aba_tambem_encerra_o_bloco():
    # Cabeçalho sem nem linha de código embaixo (fim do used range): não vira
    # bloco, mas ainda é fronteira.
    ws = _Sheet("Cargo", rows=8, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Ana", None, "OK"), ("Bruno", None, "OK")])
    ws.range(8, 1).value = "Treinamentos"
    blocos, _ = planos_macs._ler_blocos(ws)
    assert [(b["linha_cod"], b["linha_fim"]) for b in blocos] == [(2, 8)]


def test_pseudo_codigo_continua_virando_bloco():
    # "Treinamento Global\nLe@rn-ID- 00610773" tem 2 hífens: segue sendo
    # "código" (inofensivo, nunca está na Matriz). A regra não mudou.
    ws = _Sheet("Cargo", rows=12, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Ana", None, "OK")])
    _bloco(ws, 8, "Treinamento Global\nLe@rn-ID- 00610773", [("Carla", None, None)])
    blocos, _ = planos_macs._ler_blocos(ws)
    assert [(b["codigos"], b["linha_fim"]) for b in blocos] == [
        ({3: "DOC-POP-1"}, 8),
        ({3: "Treinamento Global\nLe@rn-ID- 00610773"}, 13),
    ]


def test_aba_so_com_blocos_sem_codigo_continua_travando():
    # Os cabeçalhos sem código viraram fronteira, mas NÃO contam como bloco:
    # a trava R2 ("tem dados e nenhum bloco reconhecido") segue igual.
    ws = _Sheet("Cargo", rows=7, cols=5)
    _bloco(ws, 1, "Ilearn", [("Carla", None, "OK")])
    blocos, trava = planos_macs._blocos_e_trava_da_aba(ws)
    assert blocos == []
    assert trava and "nenhum bloco 'Treinamentos'" in trava


# ---------------------------------------------------------------------------
# (a) + (c) ciclo: reconhecido -> sem código -> reconhecido
# ---------------------------------------------------------------------------
def test_ciclo_nao_marca_quem_esta_no_bloco_sem_codigo():
    ws = _aba_reconhecido_sem_codigo_reconhecido()
    codigos, marcacoes, atrasados, desconhecidos = planos_macs._atualizar_aba(
        ws, _avaliacao("DOC-POP-1"), MESES, HOJE
    )

    # (c) as pessoas do bloco reconhecido continuam marcadas normalmente.
    assert codigos == {"DOC-POP-1", "DOC-POP-2"}
    assert _linhas_marcadas(marcacoes) == [5, 6]
    assert ws.valor(5, 4) == "ON TIME"
    assert ws.valor(6, 4) == "ON TIME"
    assert "20/10/2026" in ws.valor(5, 3)

    # (a) a Carla (bloco sem código) não é tocada: o OK dela é do OUTRO curso.
    assert ws.valor(12, 3) is None
    assert ws.valor(12, 4) == "OK"
    # E o Davi (DOC-POP-2, que não versionou) também não.
    assert ws.valor(18, 4) == "OK"
    assert atrasados == []
    assert desconhecidos == []


def test_ciclo_nao_le_o_cabecalho_do_proximo_bloco_como_pessoa():
    # Antes a faixa ia até o linha_cod do próximo bloco e INCLUÍA a linha do
    # cabeçalho dele. Se a coluna B dessa linha tivesse texto, virava "pessoa".
    ws = _aba_reconhecido_sem_codigo_reconhecido()
    ws.range(8, 2).value = "Treinamentos BPF"
    ws.range(14, 2).value = "Treinamentos Sistemas"
    _c, marcacoes, _a, _d = planos_macs._atualizar_aba(
        ws, _avaliacao("DOC-POP-1", "DOC-POP-2"), MESES, HOJE
    )
    assert _linhas_marcadas(marcacoes) == [5, 6, 18]
    assert ws.valor(8, 4) is None
    assert ws.valor(14, 4) is None


def test_concluir_de_quem_so_esta_no_bloco_sem_codigo_nao_acha_nada():
    ws = _aba_reconhecido_sem_codigo_reconhecido()
    assert conclusao._localizar_na_aba(ws, "CARLA", "DOC-POP-1") == []


def test_concluir_no_bloco_reconhecido_nao_alcanca_o_bloco_sem_codigo():
    # O caso medido nos dados reais (ex.: uma analista): a pessoa está no bloco
    # reconhecido E em blocos sem código abaixo. Só a ocorrência do bloco
    # reconhecido é dela.
    ws = _aba_reconhecido_sem_codigo_reconhecido(
        pessoas_sem_codigo=[("Carla", None, "OK"), ("Ana", None, None)]
    )
    assert conclusao._localizar_na_aba(ws, "ANA", "DOC-POP-1") == [
        conclusao.OcorrenciaPlanos("Cargo", 5, 3, 2, "OK")
    ]


def test_registrar_conclusao_de_quem_so_esta_no_bloco_sem_codigo_e_erro(
    monkeypatch, tmp_path, caplog, _sem_caminho_real_nem_excel
):
    pasta_planos = _sem_caminho_real_nem_excel
    arquivo = pasta_planos / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("antes", encoding="utf-8")
    ws = _aba_reconhecido_sem_codigo_reconhecido()
    wb = _Workbook([ws])
    _usar_workbook(monkeypatch, wb, ["Cargo"])

    codigo_saida = conclusao.registrar_conclusao(
        "Carla", "DOC-POP-1", mes_planos=9, ano_planos=2026
    )

    assert codigo_saida == 1
    assert "Não encontrei 'Carla' + curso 'DOC-POP-1'" in caplog.text
    assert not wb.salvou
    assert ws.valor(12, 3) is None
    assert ws.valor(12, 4) == "OK"
    assert arquivo.read_text(encoding="utf-8") == "antes"
    # Nem a transação começou (nenhuma cópia de restauração).
    assert not (pasta_planos / "_pre-colagem").exists()


def test_registrar_conclusao_no_bloco_reconhecido_grava_so_a_celula_dela(
    monkeypatch, _sem_caminho_real_nem_excel
):
    pasta_planos = _sem_caminho_real_nem_excel
    arquivo = pasta_planos / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("antes", encoding="utf-8")
    ws = _aba_reconhecido_sem_codigo_reconhecido(
        pessoas_sem_codigo=[("Carla", None, "OK"), ("Ana", "05/10/2026", "ON TIME")]
    )
    ws.range(5, 3).value = "'20/10/2026"
    ws.range(5, 4).value = "ON TIME"
    wb = _Workbook([ws])
    _usar_workbook(monkeypatch, wb, ["Cargo"])

    assert (
        conclusao.registrar_conclusao("Ana", "DOC-POP-1", mes_planos=9, ano_planos=2026)
        == 0
    )
    assert wb.salvou
    assert ws.valor(5, 3) is None
    assert ws.valor(5, 4) == "OK"
    # A linha da Ana no bloco sem código (L13) é de OUTRO curso: intacta.
    assert ws.valor(13, 3) == "05/10/2026"
    assert ws.valor(13, 4) == "ON TIME"


# ---------------------------------------------------------------------------
# (b) última aba: bloco sem código depois do último reconhecido
# ---------------------------------------------------------------------------
def test_ciclo_completo_ultima_aba_com_bloco_sem_codigo_no_fim(
    monkeypatch, _sem_caminho_real_nem_excel
):
    pasta_planos = _sem_caminho_real_nem_excel
    arquivo = pasta_planos / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("fake", encoding="utf-8")
    aba_a = _aba_reconhecido_sem_codigo_reconhecido("Cargo A")
    aba_b = _aba_reconhecido_e_sem_codigo_no_fim("Cargo B")
    matriz = _matriz("DOC-POP-1")
    wb = _Workbook([aba_a, aba_b, matriz])
    _usar_workbook(monkeypatch, wb, ["Cargo A", "Cargo B"])

    assert planos_macs.atualizar_planos_e_macs(
        _avaliacao("DOC-POP-1"), MESES, arquivo, hoje=HOJE
    )
    assert wb.salvou

    # (c) quem está no bloco reconhecido foi marcado, nas duas abas.
    assert aba_a.valor(5, 4) == "ON TIME"
    assert aba_a.valor(6, 4) == "ON TIME"
    assert aba_b.valor(5, 4) == "ON TIME"
    # (b) o bloco sem código do fim da última aba ficou intacto: nem o OK
    # legítimo da Carla foi sobrescrito, nem a célula vazia do Edu ganhou status.
    assert aba_b.valor(12, 3) is None
    assert aba_b.valor(12, 4) == "OK"
    assert aba_b.valor(13, 3) is None
    assert aba_b.valor(13, 4) is None
    # (a) idem no bloco sem código do meio da primeira aba.
    assert aba_a.valor(12, 4) == "OK"
    # Quitação continua sendo a última mutação.
    assert matriz.valor(5, 5) == "'2.0"  # quitação grava texto protegido


def test_concluir_de_quem_so_esta_no_bloco_sem_codigo_do_fim_e_erro(
    monkeypatch, _sem_caminho_real_nem_excel
):
    pasta_planos = _sem_caminho_real_nem_excel
    arquivo = pasta_planos / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("antes", encoding="utf-8")
    aba_a = _aba_reconhecido_sem_codigo_reconhecido("Cargo A")
    aba_b = _aba_reconhecido_e_sem_codigo_no_fim("Cargo B")
    wb = _Workbook([aba_a, aba_b])
    _usar_workbook(monkeypatch, wb, ["Cargo A", "Cargo B"])

    assert conclusao._localizar_na_aba(aba_b, "EDU", "DOC-POP-1") == []
    assert (
        conclusao.registrar_conclusao("Edu", "DOC-POP-1", mes_planos=9, ano_planos=2026)
        == 1
    )
    assert not wb.salvou
    assert aba_b.valor(13, 4) is None
    assert aba_b.valor(12, 4) == "OK"
    assert arquivo.read_text(encoding="utf-8") == "antes"


# ---------------------------------------------------------------------------
# (d) ELABORADOR / ELABORADORA / APROVADOR / APROVADORA
# ---------------------------------------------------------------------------
VARIANTES_PRESERVADAS = [
    "ELABORADOR",
    "APROVADORA",
    "Elaborador",
    "Aprovadora",
    "Elaboradora ",  # com espaço no fim (caso real, aba Gerente)
    "Elaboradora",
    "elaboradora",
    "Aprovador",
    " APROVADOR ",
    "aPrOvAdOrA",
]

PARECIDOS_NAO_PRESERVADOS = [
    "Elaboração",
    "Elaborado",
    "Elaboradores",
    "Aprovação",
    "Aprovado",
    "Revisor",
]


@pytest.mark.parametrize("planejado", VARIANTES_PRESERVADAS)
def test_variantes_de_elaborador_aprovador_sao_reconhecidas(planejado):
    assert planejado_de_elaborador_ou_aprovador(planejado)
    for status in ("OK", "ok ", None, ""):
        assert decidir_marcacao_planos(status, planejado, PRAZO, HOJE) == MARCACAO_PRESERVAR


@pytest.mark.parametrize("planejado", PARECIDOS_NAO_PRESERVADOS + [None, PRAZO, 45000.0])
def test_texto_parecido_nao_e_elaborador_nem_aprovador(planejado):
    assert not planejado_de_elaborador_ou_aprovador(planejado)
    assert decidir_marcacao_planos("OK", planejado, PRAZO, HOJE) == MARCACAO_ON_TIME


@pytest.mark.parametrize("planejado", VARIANTES_PRESERVADAS)
def test_ciclo_preserva_o_texto_original_da_variante(planejado):
    ws = _Sheet("Cargo", rows=5, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Lia", planejado, "OK")])
    _c, marcacoes, _a, desconhecidos = planos_macs._atualizar_aba(
        ws, _avaliacao("DOC-POP-1"), MESES, HOJE
    )
    assert marcacoes == []
    assert desconhecidos == []
    # Texto original intacto, inclusive o espaço do fim e a caixa.
    assert ws.valor(5, 3) == planejado
    assert ws.valor(5, 4) == "OK"


@pytest.mark.parametrize("planejado", PARECIDOS_NAO_PRESERVADOS)
def test_ciclo_segue_a_regra_normal_pra_texto_parecido(planejado):
    ws = _Sheet("Cargo", rows=5, cols=5)
    _bloco(ws, 1, "DOC-POP-1", [("Lia", planejado, "OK")])
    _c, marcacoes, _a, _d = planos_macs._atualizar_aba(
        ws, _avaliacao("DOC-POP-1"), MESES, HOJE
    )
    assert _linhas_marcadas(marcacoes) == [5]
    assert ws.valor(5, 4) == "ON TIME"
    assert "20/10/2026" in ws.valor(5, 3)
