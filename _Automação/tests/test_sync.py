from datetime import datetime

import pytest

from treinamentos_its.sync import obter_meses


def test_obter_meses_padrao(monkeypatch):
    class Agora(datetime):
        @classmethod
        def now(cls):
            return cls(2026, 9, 25)

    monkeypatch.setattr("treinamentos_its.sync.datetime", Agora)
    meses = obter_meses(None, None)
    # Regra do negócio (06/10): sem escolher, o ciclo é o do MÊS PASSADO (os POPs de
    # Agosto chegam em setembro) e cria o Planos e Macs dele a partir do anterior.
    assert (meses["anterior"], meses["atual"]) == ("Julho", "Agosto")
    assert (meses["ano_ant_relativo"], meses["ano_atual"]) == (2026, 2026)


def test_obter_meses_vira_o_ano():
    # O mês informado é o dos POPs = o do arquivo CRIADO; parte do mês anterior.
    meses = obter_meses(1, 2027)
    assert (meses["anterior"], meses["ano_ant_relativo"]) == ("Dezembro", 2026)
    assert (meses["atual"], meses["ano_atual"]) == ("Janeiro", 2027)


def test_pops_de_setembro_criam_o_planos_e_macs_de_setembro_a_partir_de_agosto():
    meses = obter_meses(9, 2026)
    assert (meses["anterior"], meses["atual"]) == ("Agosto", "Setembro")


@pytest.mark.parametrize("mes", [0, 13])
def test_obter_meses_recusa_mes_invalido(mes):
    with pytest.raises(ValueError):
        obter_meses(mes, 2026)


def test_obter_meses_exige_mes_e_ano_juntos():
    with pytest.raises(ValueError):
        obter_meses(8, None)


# =============================================================================
# Cobertura recriada depois do a5dad36: `sync.avaliar_arquivo_matriz` (a leitura
# do workbook que alimenta a avaliação pura única) com o livro de mentira.
# =============================================================================

import logging

from dubles_excel import (
    FORMULA_F,
    AbaFalsa,
    ExcelFalso,
    LivroFalso,
    aba_de_cargo,
    aba_matriz,
    aba_versionamento,
    abas_de_cargo_vazias,
)
from treinamentos_its import sync
from treinamentos_its.config import ABAS_DE_CARGO

CODIGO = "DOC-POP-0000001"
ANALISTA = ABAS_DE_CARGO[0]
GERENTE = ABAS_DE_CARGO[3]
APROVADO_EM = datetime(2026, 9, 1)


def _livro(
    linhas_matriz=None,
    *,
    datas=None,
    abas_cargo=None,
    sem_abas=(),
    versionamento=True,
):
    linhas_matriz = linhas_matriz or [(CODIGO, "1.0", 2.0, FORMULA_F)]
    abas_cargo = abas_cargo or [aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))]
    nomes = tuple(aba.name for aba in abas_cargo) + tuple(sem_abas)
    abas = [aba_matriz(linhas_matriz), *abas_cargo, *abas_de_cargo_vazias(exceto=nomes)]
    if versionamento:
        abas.append(aba_versionamento({CODIGO: APROVADO_EM} if datas is None else datas))
    return LivroFalso(*abas)


@pytest.fixture
def arquivo(tmp_path):
    caminho = tmp_path / "Planos e Macs Setembro 2026.xlsx"
    caminho.write_text("marcador", encoding="utf-8")
    return caminho


def _avaliar(monkeypatch, arquivo, livro, caplog=None):
    excel = ExcelFalso(livro).instalar(monkeypatch, sync)
    if caplog is not None:
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            resultado = sync.avaliar_arquivo_matriz(arquivo)
    else:
        resultado = sync.avaliar_arquivo_matriz(arquivo)
    return resultado, excel


class TestAvaliarArquivoMatriz:
    def test_f_com_formula_le_versionamento_prazo_e_quitacao(self, monkeypatch, arquivo):
        livro = _livro()
        resultado, excel = _avaliar(monkeypatch, arquivo, livro)

        assert resultado.pode_executar
        assert resultado.alteracoes == {CODIGO: {"v_antiga": "1.0", "v_nova": "2.0"}}
        assert resultado.prazos == {CODIGO: datetime(2026, 10, 1)}
        assert resultado.linhas_quitacao == {CODIGO: 5}
        assert excel.aberturas == [(arquivo, True)]  # só leitura
        assert livro.fechamentos == 1 and livro.salvamentos == 0

    def test_f_digitada_por_cima_da_formula_trava(self, monkeypatch, arquivo, caplog):
        livro = _livro([(CODIGO, "1.0", 2.0, None)])
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro, caplog)

        assert resultado.sem_formula_f == [f"L5 {CODIGO}"]
        assert not resultado.pode_executar
        assert f"O ciclo vai PARAR por isto: F preenchida sem fórmula: L5 {CODIGO}" in caplog.text

    def test_formula_lida_linha_a_linha(self, monkeypatch, arquivo):
        livro = _livro(
            [
                (CODIGO, "1.0", 2.0, FORMULA_F),
                ("DOC-POP-0000002", "3.0", 3.0, None),  # valor colado por cima
                ("DOC-POP-0000003", "1.0", None, None),  # POP recém-cadastrado, F vazia
            ],
            datas={CODIGO: APROVADO_EM},
        )
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro)
        assert resultado.sem_formula_f == ["L6 DOC-POP-0000002"]
        assert resultado.sem_versao_nova == ["DOC-POP-0000003"]  # só aviso
        assert any(aviso.startswith("F vazia") for aviso in resultado.avisos)
        assert not any("DOC-POP-0000003" in b for b in resultado.bloqueios)

    @pytest.mark.parametrize("problema", ["aba ausente", "cabeçalho mudou"])
    def test_versionamento_ilegivel_devolve_none(self, monkeypatch, arquivo, caplog, problema):
        if problema == "aba ausente":
            livro = _livro(versionamento=False)
        else:
            livro = _livro(versionamento=False)
            aba = AbaFalsa("Versionamento Mês")
            aba.celula(1, 1).value = "Código"
            aba.celula(1, 2).value = "Data"
            aba.celula(2, 1).value = CODIGO
            aba.celula(2, 2).value = APROVADO_EM
            livro.sheets.append(aba)

        resultado, _excel = _avaliar(monkeypatch, arquivo, livro, caplog)

        assert resultado is None
        assert "não foi possível ler a aba 'Versionamento Mês'" in caplog.text
        assert livro.fechamentos == 1

    def test_aba_de_cargo_faltando_entra_nos_bloqueios(self, monkeypatch, arquivo, caplog):
        livro = _livro(sem_abas=(GERENTE,))
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro, caplog)

        assert len(resultado.travas_abas_de_cargo) == 1
        assert GERENTE in resultado.travas_abas_de_cargo[0]
        assert resultado.travas_abas_de_cargo[0] in resultado.bloqueios
        assert not resultado.pode_executar
        assert "O ciclo vai PARAR por isto: aba(s) de cargo não encontrada(s)" in caplog.text

    def test_aba_com_dados_sem_bloco_entra_nos_bloqueios(self, monkeypatch, arquivo):
        quebrada = AbaFalsa(GERENTE)
        quebrada.celula(1, 1).value = "Treinamento"  # singular: não reconhece
        quebrada.celula(5, 2).value = "Caio"
        livro = _livro(abas_cargo=[aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)), quebrada])
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro)

        assert any(GERENTE in trava for trava in resultado.travas_abas_de_cargo)
        assert not resultado.pode_executar

    def test_replay_herdado_do_rodape_trava(self, monkeypatch, arquivo):
        herdada = aba_de_cargo(
            ANALISTA,
            [("Carla", {})],
            (CODIGO,),
            rodape=("Referente Julho 2026", f"Versionado treinamento {CODIGO} de v1.0 para v2.0"),
        )
        resultado, _excel = _avaliar(monkeypatch, arquivo, _livro(abas_cargo=[herdada]))
        assert resultado.replay == [CODIGO]
        assert not resultado.pode_executar

    def test_sem_aba_matriz_devolve_none(self, monkeypatch, arquivo, caplog):
        livro = LivroFalso(*abas_de_cargo_vazias(), aba_versionamento({}))
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro, caplog)
        assert resultado is None
        assert "Aba 'Matriz - Atualização' não encontrada" in caplog.text

    def test_arquivo_inexistente_nem_abre_o_excel(self, monkeypatch, tmp_path):
        resultado, excel = _avaliar(monkeypatch, tmp_path / "nao-existe.xlsx", _livro())
        assert resultado is None
        assert excel.aberturas == []

    def test_erro_do_excel_devolve_none(self, monkeypatch, arquivo, caplog):
        def _abrir_quebrado(*_args, **_kwargs):
            raise RuntimeError("Excel não subiu")

        ExcelFalso(_livro()).instalar(monkeypatch, sync)
        monkeypatch.setattr(sync, "abrir_livro", _abrir_quebrado)
        assert sync.avaliar_arquivo_matriz(arquivo) is None
        assert "Erro ao ler a Matriz" in caplog.text

    def test_erro_nas_travas_identifica_as_abas_de_cargo(
        self, monkeypatch, arquivo, caplog
    ):
        ExcelFalso(_livro()).instalar(monkeypatch, sync)

        def _travas_quebradas(_livro):
            raise RuntimeError("aba ilegível")

        monkeypatch.setattr(sync, "travas_das_abas_de_cargo", _travas_quebradas)

        assert sync.avaliar_arquivo_matriz(arquivo) is None
        assert "Erro ao ler as abas de cargo" in caplog.text
        assert "Erro ao ler a Matriz" not in caplog.text

    def test_log_do_prazo_mostra_a_data_de_aprovacao(self, monkeypatch, arquivo, caplog):
        # Correção J do 510a997: o prazo tem que dar pra conferir na mão.
        _resultado, _excel = _avaliar(monkeypatch, arquivo, _livro(), caplog)
        assert (
            f"Versionamento {CODIGO}: v1.0 -> v2.0; prazo 01/10/2026 "
            f"(aprovado em 01/09/2026 + 30 dias)."
        ) in caplog.text

    def test_log_sem_data_de_aprovacao_diz_onde_faltou(self, monkeypatch, arquivo, caplog):
        resultado, _excel = _avaliar(monkeypatch, arquivo, _livro(datas={}), caplog)
        assert resultado.sem_data_aprovacao == [CODIGO]
        assert not resultado.pode_executar
        assert (
            f"Versionamento {CODIGO}: v1.0 -> v2.0; prazo indisponível (sem data de "
            f"aprovação na aba 'Versionamento Mês')."
        ) in caplog.text


class TestColunaEComFormula:
    """Decisão de negócio (2026-10-02): E com fórmula trava o sync e o conferir."""

    def test_e_com_formula_e_lida_e_trava(self, monkeypatch, arquivo, caplog):
        from treinamentos_its.config import COL_MATRIZ_VERSAO_ATUAL
        from treinamentos_its.sync import NOME_ABA_MATRIZ

        livro = _livro([(CODIGO, "2.0", 2.0, FORMULA_F)])
        livro.sheets[NOME_ABA_MATRIZ].celula(5, COL_MATRIZ_VERSAO_ATUAL).formula = (
            "=VLOOKUP(C5,'Versionamento Mês'!$A$1:$D$832,4,0)"
        )
        resultado, _excel = _avaliar(monkeypatch, arquivo, livro, caplog)

        assert resultado.e_com_formula == [f"L5 {CODIGO}"]
        assert not resultado.pode_executar
        assert "O ciclo vai PARAR por isto: E com fórmula" in caplog.text

    def test_e_com_valor_nao_trava(self, monkeypatch, arquivo):
        resultado, _excel = _avaliar(monkeypatch, arquivo, _livro())
        assert resultado.e_com_formula == []
        assert resultado.pode_executar
