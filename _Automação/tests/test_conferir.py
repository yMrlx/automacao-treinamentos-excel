from treinamentos_its import conferir
from treinamentos_its.dominio import AvaliacaoMatriz


MESES = {
    "anterior": "Agosto",
    "ano_ant_relativo": 2026,
    "atual": "Setembro",
    "ano_atual": 2026,
}


def test_conferir_sem_export_usa_o_anterior_como_esta(monkeypatch, tmp_path):
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    anterior = pasta / "Planos e Macs Agosto 2026.xlsx"
    anterior.write_text("x")
    vistos = []
    monkeypatch.setattr(conferir, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(conferir, "obter_meses", lambda *_: MESES)
    monkeypatch.setattr(
        conferir,
        "_analisar_e_cruzar",
        lambda arquivo, **_kwargs: vistos.append(arquivo) or True,
    )
    assert conferir.conferir_matriz(8, 2026)
    assert vistos == [anterior]


def test_conferir_com_export_parte_do_novo_se_ele_existe(monkeypatch, tmp_path):
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    (pasta / "Planos e Macs Agosto 2026.xlsx").write_text("anterior")
    novo = pasta / "Planos e Macs Setembro 2026.xlsx"
    novo.write_text("novo")
    export = tmp_path / "vigentes.xlsx"
    export.write_text("export")
    origens = []
    monkeypatch.setattr(conferir, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(conferir, "obter_meses", lambda *_: MESES)
    def criar(origem, destino):
        origens.append(origem)
        destino.write_text("copia")
        return True

    monkeypatch.setattr("treinamentos_its.planos_macs.criar_planos_do_mes", criar)
    monkeypatch.setattr("treinamentos_its.vigentes.importar_vigentes", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(conferir, "_analisar_e_cruzar", lambda *_args, **_kwargs: True)
    assert conferir.conferir_matriz(8, 2026, export)
    assert origens == [novo]


def test_analisar_sai_com_erro_quando_o_ciclo_vai_parar(monkeypatch, tmp_path):
    resultado = AvaliacaoMatriz(bloqueios=["regressão"])
    monkeypatch.setattr(conferir, "avaliar_arquivo_matriz", lambda _arquivo: resultado)
    assert not conferir._analisar_e_cruzar(tmp_path / "x.xlsx")


def test_export_importado_pula_comparacao_tautologica(monkeypatch, tmp_path):
    monkeypatch.setattr(
        conferir, "avaliar_arquivo_matriz", lambda _arquivo: AvaliacaoMatriz()
    )
    assert conferir._analisar_e_cruzar(
        tmp_path / "x.xlsx",
        pops_vigentes=tmp_path / "vigentes.xlsx",
        vigentes_ja_importado=True,
    )


def test_obsoletos_recebe_todos_os_codigos_da_coluna_c(monkeypatch, tmp_path):
    avaliacao = AvaliacaoMatriz(
        codigos_matriz=["NUMERICO", "VERIFICAR", "F-VAZIA", "F-TEXTO"],
        versoes_matriz={"NUMERICO": "2.0"},
        verificar=["VERIFICAR"],
    )
    recebidos = []
    monkeypatch.setattr(conferir, "avaliar_arquivo_matriz", lambda _arquivo: avaliacao)
    monkeypatch.setattr(
        conferir,
        "_conferir_contra_obsoletos",
        lambda codigos, _arquivo: recebidos.extend(codigos) or True,
    )

    assert conferir._analisar_e_cruzar(
        tmp_path / "matriz.xlsx", pops_obsoletos=tmp_path / "obsoletos.xlsx"
    )
    assert recebidos == ["NUMERICO", "VERIFICAR", "F-VAZIA", "F-TEXTO"]


# =============================================================================
# Cobertura das correções B e H do 510a997: o export de obsoletos achado sozinho
# e o export ilegível virando exit 1 sem traceback. Nada abre Excel.
# =============================================================================

import logging
import os
from types import SimpleNamespace

import pytest

from dubles_excel import AbaFalsa, ExcelFalso, LivroFalso
from treinamentos_its import cli, config


@pytest.fixture
def pastas(monkeypatch, tmp_path):
    pasta = tmp_path / "Planos e Macs"
    sistema = pasta / "POPs do Sistema"
    sistema.mkdir(parents=True)
    (pasta / "Planos e Macs Agosto 2026.xlsx").write_text("agosto", encoding="utf-8")
    monkeypatch.setattr(conferir, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(conferir, "obter_meses", lambda *_: MESES)
    vistos: list[dict] = []
    monkeypatch.setattr(
        conferir,
        "_analisar_e_cruzar",
        lambda arquivo, **kwargs: vistos.append({"arquivo": arquivo, **kwargs}) or True,
    )
    return SimpleNamespace(pasta=pasta, sistema=sistema, vistos=vistos)


def _export(pasta, nome, idade_em_dias=0):
    arquivo = pasta / nome
    arquivo.write_text("export", encoding="utf-8")
    momento = 1_780_000_000 - idade_em_dias * 86_400
    os.utime(arquivo, (momento, momento))
    return arquivo


class TestObsoletosAchadoSozinho:
    def test_usa_o_mais_recente_de_pops_do_sistema(self, pastas, caplog):
        _export(pastas.sistema, "POPs Obsoletos Julho 2026.xlsx", idade_em_dias=30)
        recente = _export(pastas.sistema, "POPs Obsoletos Agosto 2026.xlsx")
        _export(pastas.sistema, "POPs aprovados e efetivos Agosto 2026.xlsx")

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conferir.conferir_matriz(8, 2026)

        assert pastas.vistos[0]["pops_obsoletos"] == recente
        assert "usando o mais recente da pasta" in caplog.text
        assert str(recente) in caplog.text

    def test_cai_na_raiz_do_planos_quando_a_subpasta_esta_vazia(self, pastas):
        na_raiz = _export(pastas.pasta, "POPs Obsoletos Agosto 2026.xlsx")
        assert conferir.conferir_matriz(8, 2026)
        assert pastas.vistos[0]["pops_obsoletos"] == na_raiz

    def test_sem_nenhum_export_segue_sem_obsoletos_e_avisa(self, pastas, caplog):
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conferir.conferir_matriz(8, 2026)
        assert pastas.vistos[0]["pops_obsoletos"] is None
        assert "Nenhum export 'POPs Obsoletos ...xlsx' indicado nem encontrado" in caplog.text

    def test_indicado_vence_o_da_pasta(self, pastas, tmp_path, caplog):
        _export(pastas.sistema, "POPs Obsoletos Agosto 2026.xlsx")
        indicado = _export(tmp_path, "meu obsoletos.xlsx", idade_em_dias=99)
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conferir.conferir_matriz(8, 2026, pops_obsoletos=indicado)
        assert pastas.vistos[0]["pops_obsoletos"] == indicado
        assert "INDICADO pelo operador" in caplog.text

    def test_indicado_que_nao_existe_nao_confere(self, pastas, tmp_path, caplog):
        assert conferir.conferir_matriz(8, 2026, pops_obsoletos=tmp_path / "sumiu.xlsx") is False
        assert pastas.vistos == []
        assert "NÃO FOI POSSÍVEL CONFERIR" in caplog.text

    def test_com_vigentes_tambem_acha_o_obsoletos_sozinho(self, pastas, monkeypatch, tmp_path):
        recente = _export(pastas.sistema, "POPs Obsoletos Agosto 2026.xlsx")
        vigentes = _export(tmp_path, "vigentes.xlsx")

        def _criar(origem, destino):
            destino.write_text("copia", encoding="utf-8")
            return True

        monkeypatch.setattr("treinamentos_its.planos_macs.criar_planos_do_mes", _criar)
        monkeypatch.setattr(
            "treinamentos_its.vigentes.importar_vigentes", lambda *_a, **_k: True
        )
        assert conferir.conferir_matriz(8, 2026, vigentes)
        assert pastas.vistos[0]["pops_obsoletos"] == recente
        assert pastas.vistos[0]["vigentes_ja_importado"] is True

    def test_pastas_dos_exports_seguem_a_pasta_redirecionada(self, pastas):
        # Nunca cair na pasta REAL quando o Planos foi redirecionado (sandbox/teste).
        sistema, raiz = conferir._pastas_dos_exports()
        assert (sistema, raiz) == (pastas.sistema, pastas.pasta)
        assert config.PASTA_POPS_SISTEMA != pastas.sistema


class _AppFalso:
    """Contexto de Excel que falha antes de abrir o workbook."""

    def __init__(self, erro=None):
        self.erro = erro

    def __call__(self, **_kwargs):
        return self

    def __enter__(self):
        if self.erro is not None:
            raise self.erro
        return self

    def __exit__(self, *_args):
        return False


def _export_obsoletos(cabecalho=("Document Number", "Major Version Number")):
    aba = AbaFalsa("Export")
    for coluna, titulo in enumerate(cabecalho, start=1):
        aba.celula(1, coluna).value = titulo
    aba.celula(2, 1).value = "DOC-POP-0000009"
    aba.celula(2, 2).value = 3.0
    aba.celula(3, 1).value = "DOC-POP-0000001"
    aba.celula(3, 2).value = 1.0
    return LivroFalso(aba)


def _sem_traceback(caplog) -> bool:
    return all(registro.exc_info is None for registro in caplog.records)


class TestObsoletosIlegivel:
    def test_export_que_nao_abre_vira_erro_claro_sem_traceback(
        self, monkeypatch, tmp_path, caplog
    ):
        monkeypatch.setattr(
            conferir, "app_excel", _AppFalso(erro=OSError("arquivo corrompido"))
        )
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert conferir._conferir_contra_obsoletos(
                ["DOC-POP-0000001"], tmp_path / "POPs Obsoletos.xlsx"
            ) is False
        assert "NÃO FOI POSSÍVEL LER o export de obsoletos" in caplog.text
        assert "OSError: arquivo corrompido" in caplog.text
        assert _sem_traceback(caplog)

    def test_cabecalho_nao_reconhecido_nao_e_ok(self, monkeypatch, tmp_path, caplog):
        livro = _export_obsoletos(cabecalho=("Código", "Versão"))
        ExcelFalso(livro).instalar(monkeypatch, conferir)
        assert conferir._conferir_contra_obsoletos(["DOC-POP-0000001"], tmp_path / "o.xlsx") is False
        assert "Cabeçalho não reconhecido" in caplog.text
        assert livro.fechamentos == 1

    def test_export_legivel_aponta_codigo_obsoleto(self, monkeypatch, tmp_path, caplog):
        caminho = tmp_path / "o.xlsx"
        excel = ExcelFalso(_export_obsoletos()).instalar(monkeypatch, conferir)
        assert conferir._conferir_contra_obsoletos(
            ["DOC-POP-0000001", "DOC-POP-0000002"], caminho
        ) is True
        assert excel.aberturas == [(caminho, True)]
        assert "1 código(s) da Matriz constam" in caplog.text
        assert "DOC-POP-0000001" in caplog.text

    def test_analisar_com_obsoletos_ilegivel_e_conferencia_incompleta(
        self, monkeypatch, tmp_path, caplog
    ):
        monkeypatch.setattr(
            conferir, "avaliar_arquivo_matriz", lambda _arquivo: AvaliacaoMatriz()
        )
        monkeypatch.setattr(conferir, "app_excel", _AppFalso(erro=OSError("travado")))
        assert conferir._analisar_e_cruzar(
            tmp_path / "x.xlsx", pops_obsoletos=tmp_path / "o.xlsx"
        ) is False
        assert "Conferência INCOMPLETA" in caplog.text
        assert "Conferência concluída" not in caplog.text

    def test_pelo_main_sai_com_1_sem_traceback(self, monkeypatch, tmp_path, caplog):
        pasta = tmp_path / "Planos e Macs"
        (pasta / "POPs do Sistema").mkdir(parents=True)
        (pasta / "Planos e Macs Agosto 2026.xlsx").write_text("agosto", encoding="utf-8")
        obsoletos = _export(pasta / "POPs do Sistema", "POPs Obsoletos Agosto 2026.xlsx")
        monkeypatch.setattr(conferir, "PASTA_PLANOS_MACS", pasta)
        monkeypatch.setattr(conferir, "obter_meses", lambda *_: MESES)
        monkeypatch.setattr(
            conferir, "avaliar_arquivo_matriz", lambda _arquivo: AvaliacaoMatriz()
        )
        monkeypatch.setattr(
            conferir, "app_excel", _AppFalso(erro=OSError("corrompido"))
        )
        monkeypatch.setattr("sys.argv", ["main.py", "conferir", "--pops-obsoletos", str(obsoletos)])

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            with pytest.raises(SystemExit) as saida:
                cli.main()

        assert saida.value.code == 1
        assert _sem_traceback(caplog)
        assert "NÃO FOI POSSÍVEL LER o export de obsoletos" in caplog.text
