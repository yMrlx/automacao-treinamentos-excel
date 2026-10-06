from types import SimpleNamespace

from treinamentos_its import backup, conclusao


class _Cell:
    def __init__(self, value=None):
        self.value = value
        self.color = "cor"
        self.font = SimpleNamespace(color="fonte")


class _Sheet:
    name = "Cargo"

    def __init__(self):
        self.cells = {
            (5, 2): _Cell("Carla"),
            (2, 3): _Cell("DOC-POP-1"),
            (5, 3): _Cell("01/10/2026"),
            (5, 4): _Cell("ON TIME"),
        }
        self.used_range = SimpleNamespace(last_cell=SimpleNamespace(row=5))

    def range(self, linha, coluna):
        return self.cells.setdefault((linha, coluna), _Cell())


def test_marcar_ok_esvazia_planejado_e_grava_ok():
    ws = _Sheet()
    ocorrencia = conclusao.OcorrenciaPlanos("Cargo", 5, 3, 2)
    conclusao._marcar_ok(ws, ocorrencia)
    assert ws.range(5, 3).value is None
    assert ws.range(5, 3).color is None
    assert ws.range(5, 4).value == "OK"
    assert ws.range(5, 4).color is None


def test_localizar_na_aba_devolve_todas_as_ocorrencias():
    # Adaptado em 02/10 (fim de bloco = próximo cabeçalho 'Treinamentos'): em vez
    # de dublar o dict do bloco (que ganhou a chave 'linha_fim'), usa a aba de
    # mentira e o `mapear_codigos_da_aba` de verdade.
    from dubles_excel import aba_de_cargo

    ws = aba_de_cargo(
        "Cargo",
        [("Ana", {}), ("Carla", {"DOC-POP-1": ("'01/10/2026", "ON TIME")})],
        ("DOC-POP-1", "DOC-POP-2"),
    )
    assert conclusao._localizar_na_aba(ws, "CARLA", "DOC-POP-1") == [
        conclusao.OcorrenciaPlanos("Cargo", 6, 3, 2, "ON TIME")
    ]


def _arquivo(monkeypatch, tmp_path):
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    arquivo = pasta / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("antes", encoding="utf-8")
    monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(
        conclusao,
        "caminho_copia_pre_colagem",
        lambda _nome: tmp_path / "copias" / "antes.xlsx",
    )
    monkeypatch.setattr(
        conclusao,
        "_localizar_ocorrencias",
        lambda *_: [conclusao.OcorrenciaPlanos("Cargo", 5, 3, 2)],
    )
    return arquivo


def test_registrar_conclusao_confirma_depois_de_reabrir(monkeypatch, tmp_path):
    arquivo = _arquivo(monkeypatch, tmp_path)
    def salvar(_arquivo, ocorrencias, *_):
        arquivo.write_text("OK")
        return ocorrencias
    monkeypatch.setattr(conclusao, "_salvar_conclusao", salvar)
    monkeypatch.setattr(conclusao, "_reabrir_e_conferir", lambda *_: True)
    assert conclusao.registrar_conclusao("Carla", "DOC-POP-1", mes_planos=9, ano_planos=2026) == 0
    assert arquivo.read_text() == "OK"


def test_par_nao_encontrado_e_erro_sem_alterar(monkeypatch, tmp_path):
    arquivo = _arquivo(monkeypatch, tmp_path)
    monkeypatch.setattr(conclusao, "_localizar_ocorrencias", lambda *_: [])
    assert conclusao.registrar_conclusao("Carla", "DOC-POP-1", mes_planos=9, ano_planos=2026) == 1
    assert arquivo.read_text(encoding="utf-8") == "antes"


def test_falha_na_conferencia_restaura_o_unico_arquivo(monkeypatch, tmp_path):
    arquivo = _arquivo(monkeypatch, tmp_path)
    def salvar(_arquivo, ocorrencias, *_):
        arquivo.write_text("mudou")
        return ocorrencias
    monkeypatch.setattr(conclusao, "_salvar_conclusao", salvar)
    monkeypatch.setattr(conclusao, "_reabrir_e_conferir", lambda *_: False)
    assert conclusao.registrar_conclusao("Carla", "DOC-POP-1", mes_planos=9, ano_planos=2026) == 1
    assert arquivo.read_text(encoding="utf-8") == "antes"


def test_restauracao_incompleta_sai_com_3(monkeypatch, tmp_path):
    arquivo = _arquivo(monkeypatch, tmp_path)
    def salvar(_arquivo, ocorrencias, *_):
        arquivo.write_text("mudou")
        return ocorrencias
    monkeypatch.setattr(conclusao, "_salvar_conclusao", salvar)
    monkeypatch.setattr(conclusao, "_reabrir_e_conferir", lambda *_: False)
    monkeypatch.setattr(backup, "restaurar_arquivo", lambda *_args, **_kwargs: False)
    assert conclusao.registrar_conclusao("Carla", "DOC-POP-1", mes_planos=9, ano_planos=2026) == 3


# =============================================================================
# Cobertura recriada depois do a5dad36 (concluir só no Planos e Macs): escolha
# do mês, conferência de identidade antes de escrever, aba sumida, AutoSave,
# reabrir e conferir. Planilha de mentira de `dubles_excel`; nada abre Excel.
# =============================================================================

import logging
from datetime import datetime

import pytest

from dubles_excel import (
    ExcelFalso,
    LivroFalso,
    aba_de_cargo,
    abas_de_cargo_vazias,
    celulas_da_pessoa,
    linha_da_pessoa,
)
from treinamentos_its import config
from treinamentos_its.config import ABAS_DE_CARGO
from treinamentos_its.conclusao import _mes_ano_do_planos_macs

CODIGO = "DOC-POP-0000001"
ANALISTA = ABAS_DE_CARGO[0]
GERENTE = ABAS_DE_CARGO[3]


class TestMesDoPlanosMacs:
    def test_segue_a_data_de_conclusao_sem_mes_ano(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), None, None) == (5, 2026)

    def test_data_no_fim_do_mes(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 1, 31), None, None) == (1, 2026)

    def test_mes_ano_explicitos_mandam(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 7, 2026) == (7, 2026)
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 12, 2025) == (12, 2025)

    def test_recusa_par_incompleto(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 7, None) is None
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), None, 2026) is None

    def test_recusa_mes_fora_do_intervalo(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 0, 2026) is None
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 13, 2026) is None

    def test_aceita_extremos_do_intervalo(self):
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 1, 2026) == (1, 2026)
        assert _mes_ano_do_planos_macs(datetime(2026, 5, 15), 12, 2026) == (12, 2026)

    @pytest.mark.parametrize(
        "data, mes, ano, esperado",
        [
            (datetime(2026, 5, 15), None, None, "Planos e Macs Maio 2026.xlsx"),
            (datetime(2026, 5, 15), 7, 2026, "Planos e Macs Julho 2026.xlsx"),
        ],
        ids=["mes da --data", "--mes/--ano"],
    )
    def test_registrar_abre_o_arquivo_do_mes_escolhido(
        self, monkeypatch, tmp_path, data, mes, ano, esperado
    ):
        pasta = tmp_path / "Planos e Macs"
        pasta.mkdir()
        for nome in ("Planos e Macs Maio 2026.xlsx", "Planos e Macs Julho 2026.xlsx"):
            (pasta / nome).write_text("x", encoding="utf-8")
        monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", pasta)
        abertos: list = []
        monkeypatch.setattr(
            conclusao, "_localizar_ocorrencias", lambda arquivo, *_: abertos.append(arquivo) or []
        )
        assert conclusao.registrar_conclusao("Carla", CODIGO, data, mes, ano) == 1
        assert [arquivo.name for arquivo in abertos] == [esperado]

    @pytest.mark.parametrize("mes, ano", [(7, None), (None, 2026), (13, 2026)])
    def test_registrar_com_mes_ruim_nem_procura(self, monkeypatch, tmp_path, mes, ano):
        monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", tmp_path)
        monkeypatch.setattr(
            conclusao, "_localizar_ocorrencias", lambda *_: pytest.fail("não devia procurar")
        )
        assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 5, 15), mes, ano) == 1

    def test_registrar_sem_o_arquivo_do_mes_e_erro(self, monkeypatch, tmp_path, caplog):
        monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", tmp_path)
        monkeypatch.setattr(
            conclusao, "_localizar_ocorrencias", lambda *_: pytest.fail("não devia procurar")
        )
        assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 5, 15)) == 1
        assert "Planos e Macs do mês não encontrado" in caplog.text


def _livro_com_bia(*, em_duas_abas=False) -> LivroFalso:
    analista = aba_de_cargo(
        ANALISTA,
        [("Ana", {}), ("Carla", {CODIGO: ("'01/10/2026", "ON TIME")})],
        (CODIGO, "DOC-POP-0000002"),
    )
    abas = [analista]
    if em_duas_abas:
        abas.append(
            aba_de_cargo(GERENTE, [("carla ", {CODIGO: ("'01/08/2026", "ATRASADO")})], (CODIGO,))
        )
    nomes = tuple(aba.name for aba in abas)
    return LivroFalso(*abas, *abas_de_cargo_vazias(exceto=nomes))


def _ocorrencia(livro, aba=ANALISTA, nome="Carla"):
    return conclusao.OcorrenciaPlanos(aba, linha_da_pessoa(livro.sheets[aba], nome), 3, 2)


class TestLocalizarOcorrencias:
    def test_acha_em_todas_as_abas_somente_leitura(self, monkeypatch, tmp_path):
        livro = _livro_com_bia(em_duas_abas=True)
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)
        arquivo = tmp_path / "Planos.xlsx"

        ocorrencias = conclusao._localizar_ocorrencias(arquivo, "CARLA", CODIGO)

        assert [(o.aba, o.coluna_planejado) for o in ocorrencias] == [
            (ANALISTA, 3),
            (GERENTE, 3),
        ]
        assert excel.aberturas == [(arquivo, True)]
        assert livro.fechamentos == 1

    def test_so_o_curso_pedido(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        assert conclusao._localizar_ocorrencias(tmp_path / "x.xlsx", "CARLA", "DOC-POP-9") == []
        [ocorrencia] = conclusao._localizar_ocorrencias(
            tmp_path / "x.xlsx", "CARLA", "DOC-POP-0000002"
        )
        assert ocorrencia.coluna_planejado == 5

    def test_aba_de_cargo_ausente_e_pulada(self, monkeypatch, tmp_path):
        livro = LivroFalso(
            aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))
        )  # só uma das abas de cargo
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        assert len(conclusao._localizar_ocorrencias(tmp_path / "x.xlsx", "CARLA", CODIGO)) == 1


class TestIdentidadeAntesDeEscrever:
    def test_mesma_pessoa_e_curso_grava_ok_e_esvazia_o_planejado(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ExcelFalso(livro).instalar(monkeypatch, conclusao)

        conclusao._salvar_conclusao(tmp_path / "x.xlsx", [_ocorrencia(livro)], "CARLA", CODIGO)

        planejado, executado = celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")
        assert planejado.value is None and planejado.color is None
        assert executado.value == "OK"
        assert livro.salvamentos == 1 and livro.fechamentos == 1

    @pytest.mark.parametrize(
        "mudanca", ["outra pessoa na linha", "outro curso na coluna"]
    )
    def test_planilha_que_mudou_entre_ler_e_escrever_aborta_sem_tocar(
        self, monkeypatch, tmp_path, mudanca
    ):
        livro = _livro_com_bia()
        ocorrencia = _ocorrencia(livro)
        aba = livro.sheets[ANALISTA]
        if mudanca == "outra pessoa na linha":
            aba.celula(ocorrencia.linha, 2).value = "Caio"
        else:
            aba.celula(ocorrencia.linha_codigo, ocorrencia.coluna_planejado).value = "DOC-POP-9"
        ExcelFalso(livro).instalar(monkeypatch, conclusao)

        with pytest.raises(conclusao.PlanilhaMudouDuranteAConclusao, match="mudou"):
            conclusao._salvar_conclusao(tmp_path / "x.xlsx", [ocorrencia], "CARLA", CODIGO)

        assert aba.celula(ocorrencia.linha, 4).value == "ON TIME"
        assert livro.salvamentos == 0 and livro.fechamentos == 1

    def test_aba_sumida_aborta(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ocorrencia = conclusao.OcorrenciaPlanos("Aba Que Sumiu", 6, 3, 2)
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        with pytest.raises(conclusao.PlanilhaMudouDuranteAConclusao, match="não existe mais"):
            conclusao._salvar_conclusao(tmp_path / "x.xlsx", [ocorrencia], "CARLA", CODIGO)
        assert livro.salvamentos == 0

    def test_uma_ocorrencia_divergente_impede_todas(self, monkeypatch, tmp_path):
        livro = _livro_com_bia(em_duas_abas=True)
        boa = _ocorrencia(livro)
        ruim = _ocorrencia(livro, GERENTE, "carla ")
        livro.sheets[GERENTE].celula(ruim.linha, 2).value = "Caio"
        ExcelFalso(livro).instalar(monkeypatch, conclusao)

        with pytest.raises(conclusao.PlanilhaMudouDuranteAConclusao):
            conclusao._salvar_conclusao(tmp_path / "x.xlsx", [boa, ruim], "CARLA", CODIGO)

        assert celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")[1].value == "ON TIME"

    def test_desliga_o_autosave_antes_de_conferir_e_escrever(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        ordem: list[str] = []
        desligar_de_verdade = conclusao.desligar_autosave
        marcar_de_verdade = conclusao._marcar_ok
        identidade_de_verdade = conclusao._identidade_confere

        def _desligar(wb, *, descricao):
            ordem.append("autosave")
            return desligar_de_verdade(wb, descricao=descricao)

        monkeypatch.setattr(conclusao, "desligar_autosave", _desligar)
        monkeypatch.setattr(
            conclusao,
            "_identidade_confere",
            lambda *a: ordem.append("conferir") or identidade_de_verdade(*a),
        )
        monkeypatch.setattr(
            conclusao, "_marcar_ok", lambda *a: ordem.append("alterar") or marcar_de_verdade(*a)
        )

        conclusao._salvar_conclusao(tmp_path / "x.xlsx", [_ocorrencia(livro)], "CARLA", CODIGO)

        assert ordem == ["autosave", "conferir", "alterar"]
        assert livro.api.AutoSaveOn is False


class TestReabrirEConferirConclusao:
    def _conferir(self, monkeypatch, tmp_path, livro, ocorrencias):
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)
        resultado = conclusao._reabrir_e_conferir(tmp_path / "x.xlsx", ocorrencias, "CARLA", CODIGO)
        assert excel.aberturas == [(tmp_path / "x.xlsx", True)]
        assert livro.fechamentos == 1
        return resultado

    @pytest.mark.parametrize("planejado_salvo", [None, "", "   "])
    def test_ok_e_planejado_vazio_confere(self, monkeypatch, tmp_path, planejado_salvo):
        livro = _livro_com_bia()
        planejado, executado = celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")
        planejado.value, executado.value = planejado_salvo, "OK"
        assert self._conferir(monkeypatch, tmp_path, livro, [_ocorrencia(livro)]) is True

    def test_status_que_nao_ficou_ok(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        planejado, _executado = celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")
        planejado.value = None  # Planejado limpo, mas o status continua ON TIME
        assert self._conferir(monkeypatch, tmp_path, livro, [_ocorrencia(livro)]) is False

    def test_planejado_que_nao_ficou_vazio(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        _planejado, executado = celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")
        executado.value = "OK"  # status OK, mas a data continua lá
        assert self._conferir(monkeypatch, tmp_path, livro, [_ocorrencia(livro)]) is False

    def test_aba_sumiu(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ocorrencia = conclusao.OcorrenciaPlanos("Aba Que Sumiu", 6, 3, 2)
        assert self._conferir(monkeypatch, tmp_path, livro, [ocorrencia]) is False

    def test_outra_pessoa_na_coordenada(self, monkeypatch, tmp_path):
        livro = _livro_com_bia()
        ocorrencia = _ocorrencia(livro)
        livro.sheets[ANALISTA].celula(ocorrencia.linha, 2).value = "Caio"
        livro.sheets[ANALISTA].celula(ocorrencia.linha, 3).value = None
        livro.sheets[ANALISTA].celula(ocorrencia.linha, 4).value = "OK"
        assert self._conferir(monkeypatch, tmp_path, livro, [ocorrencia]) is False


@pytest.fixture
def pasta_do_mes(monkeypatch, tmp_path):
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    arquivo = pasta / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("planos ANTES", encoding="utf-8")
    monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(config, "PASTA_PRE_COLAGEM", pasta / "_pre-colagem")
    return arquivo


class TestRegistrarConclusaoFluxo:
    def test_conclui_em_todas_as_ocorrencias_e_confirma(self, monkeypatch, pasta_do_mes, caplog):
        livro = _livro_com_bia(em_duas_abas=True)
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao(
                "Carla", CODIGO, datetime(2026, 9, 10)
            ) == 0

        for aba, nome in ((ANALISTA, "Carla"), (GERENTE, "carla ")):
            planejado, executado = celulas_da_pessoa(livro.sheets[aba], nome)
            assert (planejado.value, executado.value) == (None, "OK"), aba
        assert [somente for _c, somente in excel.aberturas] == [True, False, True]
        assert "Conclusão confirmada" in caplog.text and "2 ocorrência(s)" in caplog.text
        assert list((pasta_do_mes.parent / "_pre-colagem").glob("*.xlsx"))

    def test_planilha_que_mudou_restaura_e_sai_1_sem_traceback(
        self, monkeypatch, pasta_do_mes, caplog
    ):
        livro = _livro_com_bia()

        class _OutraPessoaNaHoraDeEscrever(ExcelFalso):
            def abrir_livro(self, app, caminho, *, read_only=False):
                if not read_only:  # entre a leitura e a escrita, alguém mexeu
                    aba = self.livro.sheets[ANALISTA]
                    aba.celula(linha_da_pessoa(aba, "Carla"), 2).value = "Caio"
                    pasta_do_mes.write_text("AutoSave parcial", encoding="utf-8")
                return super().abrir_livro(app, caminho, read_only=read_only)

        _OutraPessoaNaHoraDeEscrever(livro).instalar(monkeypatch, conclusao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 1

        assert pasta_do_mes.read_text(encoding="utf-8") == "planos ANTES"
        registro = next(r for r in caplog.records if "mudou entre a leitura" in r.getMessage())
        assert registro.exc_info is None  # erro esperado: sem traceback assustador
        assert "Conclusão confirmada" not in caplog.text


    def test_reabertura_que_nao_confere_restaura(self, monkeypatch, pasta_do_mes, caplog):
        escrito = _livro_com_bia()
        no_disco = _livro_com_bia()  # a escrita "não ficou": continua ON TIME

        class _ExcelQuePerde(ExcelFalso):
            def __init__(self):
                super().__init__(escrito)
                self.leituras = 0

            def abrir_livro(self, app, caminho, *, read_only=False):
                self.aberturas.append((caminho, read_only))
                if read_only:
                    self.leituras += 1
                    return escrito if self.leituras == 1 else no_disco
                pasta_do_mes.write_text("AutoSave parcial", encoding="utf-8")
                return escrito

        _ExcelQuePerde().instalar(monkeypatch, conclusao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 1

        assert "não permaneceu salva depois de reabrir" in caplog.text
        assert pasta_do_mes.read_text(encoding="utf-8") == "planos ANTES"
        assert "Conclusão confirmada" not in caplog.text


class TestConcluirSomentePendentes:
    @pytest.mark.parametrize("status", ["ON TIME", " on time ", "ATRASADO", " atrasado ", None, "", " \t "])
    def test_pendente_grava_e_confirma(self, monkeypatch, pasta_do_mes, caplog, status):
        livro = _livro_com_bia()
        planejado, executado = celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")
        executado.value = status
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        assert (planejado.value, executado.value) == (None, "OK")
        assert [ro for _arquivo, ro in excel.aberturas] == [True, False, True]
        assert livro.salvamentos == 1
        assert "Conclusão confirmada" in caplog.text and "1 ocorrência(s)" in caplog.text
        assert "Nada foi alterado" not in caplog.text

    @pytest.mark.parametrize("status", [
        "OK", " ok ", "HSE", " hse ", "OUTROS", " outros ", "GLOBAL", " global ",
        "NA", " na ", "N/A", " n/a ", "DESCONHECIDO", " on  time ", 7, 0, False,
    ])
    def test_nao_pendente_fica_intacto_sem_save_nem_backup(
        self, monkeypatch, pasta_do_mes, caplog, status
    ):
        from treinamentos_its.plano_execucao import (
            ACAO_CONCLUIR, NIVEL_NEUTRO, TEXTO_NADA_ALTERADO, desfecho_da_execucao,
        )

        livro = _livro_com_bia()
        aba = livro.sheets[ANALISTA]
        planejado, executado = celulas_da_pessoa(aba, "Carla")
        executado.value = status
        planejado.color, executado.color = (1, 2, 3), (4, 5, 6)
        planejado.font.color, executado.font.color = (7, 8, 9), (10, 11, 12)
        antes = aba.estado()
        bytes_antes = pasta_do_mes.read_bytes()
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        assert aba.estado() == antes
        assert pasta_do_mes.read_bytes() == bytes_antes
        assert excel.aberturas == [(pasta_do_mes, True)]
        assert livro.salvamentos == 0 and livro.fechamentos == 1
        assert not (pasta_do_mes.parent / "_pre-colagem").exists()
        assert "Nada foi alterado" in caplog.text
        assert "Conclusão confirmada" not in caplog.text
        assert desfecho_da_execucao(ACAO_CONCLUIR, 0, caplog.text.splitlines()) == (
            TEXTO_NADA_ALTERADO, NIVEL_NEUTRO,
        )
        if conclusao.normalizar_status(status) == "OK":
            assert "já estava OK" in caplog.text
            assert not any(r.levelno >= logging.WARNING for r in caplog.records)
        else:
            aviso = next(r for r in caplog.records if r.levelno == logging.WARNING)
            texto = aviso.getMessage()
            assert pasta_do_mes.name in texto and ANALISTA in texto
            assert "L6 C4 (Executado)" in texto and "CARLA" in texto and CODIGO in texto
            assert repr(status) in texto and "não é pendente" in texto
            protegido = conclusao.normalizar_status(status) in {"HSE", "OUTROS", "GLOBAL", "NA", "N/A"}
            assert ("status protegido" if protegido else "status desconhecido") in texto

    @pytest.mark.parametrize("status_preservado", ["OK", "HSE", "OUTROS", "GLOBAL", "NA", "N/A", "NOVO"])
    def test_misto_so_confere_as_escritas_e_preserva_outra_aba(
        self, monkeypatch, pasta_do_mes, caplog, status_preservado
    ):
        livro = _livro_com_bia(em_duas_abas=True)
        aba_protegida = livro.sheets[GERENTE]
        celulas_da_pessoa(aba_protegida, "carla ")[1].value = status_preservado
        antes = aba_protegida.estado()
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        reabrir_de_verdade = conclusao._reabrir_e_conferir
        conferidas = []

        def conferir(arquivo, ocorrencias, *identidade):
            conferidas.extend(ocorrencias)
            return reabrir_de_verdade(arquivo, ocorrencias, *identidade)

        monkeypatch.setattr(conclusao, "_reabrir_e_conferir", conferir)
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        assert aba_protegida.estado() == antes
        assert celulas_da_pessoa(livro.sheets[ANALISTA], "Carla")[1].value == "OK"
        assert len(conferidas) == 1 and conferidas[0].aba == ANALISTA
        assert "Conclusão confirmada" in caplog.text and "1 ocorrência(s)" in caplog.text
        assert livro.salvamentos == 1

    def test_repetir_conclusao_nao_regrava(self, monkeypatch, pasta_do_mes, caplog):
        livro = _livro_com_bia()
        excel = ExcelFalso(livro).instalar(monkeypatch, conclusao)
        assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0
        estado_antes = livro.sheets[ANALISTA].estado()
        caplog.clear()
        excel.aberturas.clear()

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        assert livro.sheets[ANALISTA].estado() == estado_antes
        assert livro.salvamentos == 1  # só a primeira execução
        assert excel.aberturas == [(pasta_do_mes, True)]
        assert "já estava OK" in caplog.text and "Nada foi alterado" in caplog.text
        assert "Conclusão confirmada" not in caplog.text

    def test_par_nao_encontrado_continua_erro(self, monkeypatch, pasta_do_mes, caplog):
        livro = _livro_com_bia()
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        assert conclusao.registrar_conclusao("Inexistente", CODIGO, datetime(2026, 9, 10)) == 1
        assert "Não encontrei" in caplog.text
        assert "Nada foi alterado" not in caplog.text and "Conclusão confirmada" not in caplog.text
        assert livro.salvamentos == 0

    @pytest.mark.parametrize("status_novo", ["OK", "HSE", "OUTROS", "GLOBAL", "NA", "N/A", "NOVO"])
    @pytest.mark.parametrize("outra_pendente", [False, True])
    def test_reconfere_status_antes_de_escrever(
        self, monkeypatch, pasta_do_mes, caplog, status_novo, outra_pendente
    ):
        livro = _livro_com_bia(em_duas_abas=outra_pendente)
        aba = livro.sheets[ANALISTA]
        estado_externo = []

        class ExcelComEdicaoExterna(ExcelFalso):
            def abrir_livro(self, app, caminho, *, read_only=False):
                if not read_only:
                    celulas_da_pessoa(aba, "Carla")[1].value = status_novo
                    estado_externo.append(aba.estado())
                    # Simula alteração externa depois da cópia de restauração.
                    pasta_do_mes.write_text("edicao externa", encoding="utf-8")
                return super().abrir_livro(app, caminho, read_only=read_only)

        excel = ExcelComEdicaoExterna(livro).instalar(monkeypatch, conclusao)
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        assert aba.estado() == estado_externo[0]
        assert pasta_do_mes.read_text(encoding="utf-8") == "edicao externa"
        assert livro.salvamentos == int(outra_pendente)
        if outra_pendente:
            assert celulas_da_pessoa(livro.sheets[GERENTE], "carla ")[1].value == "OK"
            assert "Conclusão confirmada" in caplog.text and "1 ocorrência(s)" in caplog.text
            assert [ro for _a, ro in excel.aberturas] == [True, False, True]
        else:
            assert "Conclusão confirmada" not in caplog.text
            assert "Nada foi alterado" in caplog.text
            assert [ro for _a, ro in excel.aberturas] == [True, False]

    def test_usa_normalizacao_compartilhada_do_dominio(self):
        from treinamentos_its import dominio

        assert conclusao.normalizar_status is dominio.normalizar_status

    def test_mesmo_curso_em_duas_colunas_nao_regrava_hse(self, monkeypatch, pasta_do_mes, caplog):
        aba = aba_de_cargo(
            ANALISTA, [("Carla", {CODIGO: ("prazo", "ON TIME")})],
            (CODIGO, CODIGO, "DOC-POP-OUTRO"),
        )
        linha = linha_da_pessoa(aba, "Carla")
        aba.celula(linha, 6).value = "HSE"
        aba.celula(linha, 8).value = "ON TIME"  # outro curso não solicitado
        livro = LivroFalso(aba)
        ExcelFalso(livro).instalar(monkeypatch, conclusao)
        antes = aba.estado()
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conclusao.registrar_conclusao("Carla", CODIGO, datetime(2026, 9, 10)) == 0

        depois = aba.estado()
        assert (linha, 3) not in depois
        assert aba.celula(linha, 4).value == "OK"
        assert {c: v for c, v in depois.items() if c not in {(linha, 3), (linha, 4)}} == {
            c: v for c, v in antes.items() if c not in {(linha, 3), (linha, 4)}
        }
        assert f"L{linha} C6 (Executado)" in caplog.text
        assert "1 ocorrência(s)" in caplog.text
