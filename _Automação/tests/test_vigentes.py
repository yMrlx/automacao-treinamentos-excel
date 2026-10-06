"""Testes da parte PURA da importação dos POPs vigentes (sem Excel)."""

from pathlib import Path

import pytest

from treinamentos_its import vigentes as mod
from treinamentos_its.config import MARCADOR_VERIFICAR
from treinamentos_its.vigentes import (
    COLUNAS_CANONICAS,
    COLUNA_PROCV_CODIGO,
    COLUNA_PROCV_VERSAO,
    analisar_duplicatas_do_export,
    descrever_mes_inferido,
    divergencias_da_coluna_f,
    inferir_mes_ano,
    mapear_colunas_export,
    montar_bloco_canonico,
    valor_esperado_da_coluna_f,
    versoes_para_procv,
)

# Cabeçalhos reais, conferidos nos arquivos de 2026 (ver docstring do módulo).
CABECALHO_EN = [
    "Document Number", "Title", "Status", "Version", "Owning Department",
    "Impacted Departments", "Approved Date", "Approver", "Owner", "Reviewer",
]
CABECALHO_PT = [
    "Número do documento", "Título", "Status", "Versão", "Owning Department",
    "Impacted Departments", "Data de aprovação", "Aprovador",
]
# Junho/2026: DUAS colunas "Version" (D='9.0', E='9') e a data cai na H.
CABECALHO_JUNHO = [
    "Document Number", "Title", "Status", "Version", "Version",
    "Owning Department", "Impacted Departments", "Approved Date", "Approver",
]


def test_copia_de_restauracao_pode_ficar_na_pasta_descartavel(monkeypatch, tmp_path):
    arquivo = tmp_path / "trabalho" / "Planos e Macs Setembro 2026.xlsx"
    destino = tmp_path / "trabalho" / "_pre-colagem" / arquivo.name
    observado = {}

    def copiar(origem, caminho_copia, *, descricao):
        observado.update(origem=origem, destino=caminho_copia, descricao=descricao)
        return Path(caminho_copia)

    monkeypatch.setattr(mod, "copiar_para_restauracao", copiar)

    assert mod._salvar_copia_pre_colagem(arquivo, destino) == destino
    assert observado["origem"] == arquivo
    assert observado["destino"] == destino


class TestContratoComAFormula:
    """A fórmula da Matriz lê ``$A:$D`` com índice 4 — isso é um contrato."""

    def test_a_versao_e_a_quarta_coluna_canonica(self):
        rotulos = [rotulo for rotulo, _aceitos, _obrig in COLUNAS_CANONICAS]
        assert rotulos[COLUNA_PROCV_CODIGO - 1] == "Document Number"
        assert rotulos[COLUNA_PROCV_VERSAO - 1] == "Version"

    def test_codigo_e_versao_caem_nas_posicoes_da_formula(self):
        bloco, _avisos = montar_bloco_canonico(
            CABECALHO_JUNHO,
            [["DOC-POP-1", "T", "Efetivo", "9.0", "9", "Dep", "Imp", "01/06/2026", "Ana"]],
        )
        linha = bloco[1]
        assert linha[COLUNA_PROCV_CODIGO - 1] == "DOC-POP-1"
        assert linha[COLUNA_PROCV_VERSAO - 1] == "9.0"


class TestInferirMesAno:
    @pytest.mark.parametrize(
        "nome, esperado",
        [
            ("POPs aprovados e efetivos - Agosto 2026..xlsx", (8, 2026)),
            ("POPs Obsoletos - Julho 2026.xlsx", (7, 2026)),
            ("export marco 2026.xlsx", (3, 2026)),
            ("export Março 2026.xlsx", (3, 2026)),
            ("POPs 08-2026.xlsx", (8, 2026)),
            ("POPs 2026-08.xlsx", (8, 2026)),
            ("vigentes.xlsx", (None, None)),
            ("", (None, None)),
        ],
    )
    def test_infere_do_nome(self, nome, esperado):
        assert inferir_mes_ano(nome) == esperado

    def test_mes_sem_ano(self):
        assert inferir_mes_ano("POPs de Setembro.xlsx") == (9, None)

    def test_descricao_avisa_quando_nao_identificou(self):
        assert "NÃO identificado" in descrever_mes_inferido(None, None)
        assert "Agosto/2026" in descrever_mes_inferido(8, 2026)
        assert "ano não identificado" in descrever_mes_inferido(8, None)


class TestMapearColunas:
    def test_cabecalho_em_ingles(self):
        mapa, faltando = mapear_colunas_export(CABECALHO_EN)
        assert faltando == []
        assert mapa["Document Number"] == 1
        assert mapa["Version"] == 4
        assert mapa["Approved Date"] == 7

    def test_cabecalho_em_portugues(self):
        mapa, faltando = mapear_colunas_export(CABECALHO_PT)
        assert faltando == []
        assert mapa["Document Number"] == 1
        assert mapa["Version"] == 4
        assert mapa["Approved Date"] == 7

    def test_junho_duas_colunas_version_vence_a_da_esquerda(self):
        # A da esquerda traz "9.0", no mesmo formato da coluna E da Matriz.
        mapa, faltando = mapear_colunas_export(CABECALHO_JUNHO)
        assert faltando == []
        assert mapa["Version"] == 4
        assert mapa["Approved Date"] == 8

    def test_versao_sem_cabecalho_reconhecivel_e_obrigatoria(self):
        # Caso real: o arquivo de Abril/2026 tinha 'Coluna1' no lugar de 'Version'.
        cabecalho = list(CABECALHO_EN)
        cabecalho[3] = "Coluna1"
        _mapa, faltando = mapear_colunas_export(cabecalho)
        assert len(faltando) == 1 and "Version" in faltando[0]


class TestMontarBlocoCanonico:
    def test_junho_traz_a_data_pra_coluna_canonica_e_guarda_a_version_extra(self):
        # Em Junho/2026 a "Approved Date" fica na H (por causa da segunda coluna
        # "Version"); o código e a primeira versão já estão em A e D.
        bloco, avisos = montar_bloco_canonico(
            CABECALHO_JUNHO,
            [["POP-1", "T", "Efetivo", "9.0", "9", "Dep", "Imp", "01/06/2026", "Ana"]],
        )
        assert bloco[0][:8] == [
            "Document Number", "Title", "Status", "Version", "Owning Department",
            "Impacted Departments", "Approved Date", "Approver",
        ]
        assert bloco[1][6] == "01/06/2026"  # a data voltou pra coluna G
        # A segunda coluna "Version" do export não some: vai pro fim.
        assert bloco[0][8] == "Version"
        assert bloco[1][8] == "9"
        # Código em A e versão em D desde a origem: nada foi reposicionado.
        assert not any("REPOSICIONAD" in aviso for aviso in avisos)

    def test_versao_fora_da_quarta_coluna_e_reposicionada_com_aviso(self):
        # É o caso que quebraria o PROCV em silêncio: a fórmula lê a 4ª coluna
        # fixa, então a versão TEM que cair nela.
        cabecalho = ["Document Number", "Version", "Title", "Status", "Approved Date"]
        bloco, avisos = montar_bloco_canonico(
            cabecalho, [["POP-1", "7.0", "T", "Efetivo", "01/08/2026"]]
        )
        assert bloco[1][COLUNA_PROCV_CODIGO - 1] == "POP-1"
        assert bloco[1][COLUNA_PROCV_VERSAO - 1] == "7.0"
        assert any("REPOSICIONAD" in aviso for aviso in avisos)

    def test_preserva_o_texto_original_do_cabecalho(self):
        bloco, _avisos = montar_bloco_canonico(CABECALHO_PT, [["POP-1", "T", "E", "3", "", "", "01/02/2026", ""]])
        assert bloco[0][0] == "Número do documento"
        assert bloco[0][3] == "Versão"

    def test_descarta_linha_sem_codigo(self):
        bloco, _avisos = montar_bloco_canonico(
            CABECALHO_EN,
            [
                ["POP-1", "T", "E", "1.0", "", "", "01/08/2026", ""],
                [None, None, None, None, None, None, None, None],
                ["   ", "", "", "", "", "", "", ""],
            ],
        )
        assert len(bloco) == 2

    def test_coluna_opcional_ausente_vai_vazia_e_avisa(self):
        cabecalho = ["Document Number", "Version", "Approved Date"]
        bloco, avisos = montar_bloco_canonico(cabecalho, [["POP-1", "2.0", "01/08/2026"]])
        assert bloco[0][1] == "Title"
        assert bloco[1][1] is None
        assert bloco[1][COLUNA_PROCV_VERSAO - 1] == "2.0"
        assert any("Title" in aviso for aviso in avisos)

    def test_sem_reposicionamento_nao_avisa_a_toa(self):
        _bloco, avisos = montar_bloco_canonico(
            CABECALHO_EN, [["POP-1", "T", "E", "1.0", "D", "I", "01/08/2026", "A", "", ""]]
        )
        assert not any("REPOSICIONAD" in aviso for aviso in avisos)


class TestProcvRefeitoNoPython:
    def _bloco(self):
        return montar_bloco_canonico(
            CABECALHO_EN,
            [
                ["POP-1", "T", "E", "9.0", "", "", "01/08/2026", ""],
                ["POP-2", "T", "E", "2.0", "", "", "01/08/2026", ""],
                ["POP-1", "T", "E", "99.0", "", "", "01/08/2026", ""],
            ],
        )[0]

    def test_primeira_ocorrencia_vence_como_no_procv(self):
        versoes = versoes_para_procv(self._bloco())
        assert versoes["POP-1"] == "9.0"

    def test_lookup_ignora_caixa(self):
        versoes = versoes_para_procv(self._bloco())
        assert valor_esperado_da_coluna_f("pop-2", versoes) == "2.0"

    def test_codigo_ausente_vira_verificar(self):
        versoes = versoes_para_procv(self._bloco())
        assert valor_esperado_da_coluna_f("POP-999", versoes) == MARCADOR_VERIFICAR

    def test_codigo_vazio_nao_vira_verificar(self):
        assert valor_esperado_da_coluna_f("", {}) == ""


class TestDuplicatasDoExport:
    def test_sem_duplicata(self):
        bloco, _ = montar_bloco_canonico(
            CABECALHO_EN,
            [
                ["POP-1", "T", "E", "1.0", "", "", "01/08/2026", ""],
                ["POP-2", "T", "E", "2.0", "", "", "01/08/2026", ""],
            ],
        )
        assert analisar_duplicatas_do_export(bloco) == ([], [])

    def test_duplicata_com_mesma_versao_so_avisa(self):
        bloco, _ = montar_bloco_canonico(
            CABECALHO_EN,
            [
                ["POP-1", "T", "E", "3.0", "", "", "01/08/2026", ""],
                ["pop-1", "T", "E", 3, "", "", "02/08/2026", ""],
            ],
        )
        assert analisar_duplicatas_do_export(bloco) == (["POP-1"], [])

    def test_duplicata_com_versoes_diferentes_bloqueia(self):
        bloco, _ = montar_bloco_canonico(
            CABECALHO_EN,
            [
                ["POP-1", "T", "E", "3.0", "", "", "01/08/2026", ""],
                ["POP-1", "T", "E", "4.0", "", "", "02/08/2026", ""],
            ],
        )
        duplicados, divergentes = analisar_duplicatas_do_export(bloco)
        assert duplicados == ["POP-1"]
        assert divergentes == [("POP-1", ["3.0", "4.0"])]


class TestDivergenciasDaColunaF:
    def _versoes(self):
        return {"POP-1": "9.0", "POP-2": 2}

    def test_tudo_batendo_nao_reclama(self):
        linhas = [
            (5, "POP-1", "9.0", "=IFERROR(VLOOKUP(C5,...),\"VERIFICAR\")"),
            (6, "POP-2", 2, "=IFERROR(VLOOKUP(C6,...),\"VERIFICAR\")"),
            (7, "POP-9", "VERIFICAR", "=IFERROR(VLOOKUP(C7,...),\"VERIFICAR\")"),
        ]
        divergencias, sem_formula = divergencias_da_coluna_f(linhas, self._versoes())
        assert divergencias == []
        assert sem_formula == []

    def test_valor_velho_com_formula_e_divergencia(self):
        linhas = [(5, "POP-1", "8.0", "=IFERROR(VLOOKUP(C5,...),\"x\")")]
        divergencias, _sem = divergencias_da_coluna_f(linhas, self._versoes())
        assert len(divergencias) == 1
        assert "8.0" in divergencias[0] and "9.0" in divergencias[0]

    def test_celula_digitada_na_mao_e_aviso_nao_erro(self):
        linhas = [(5, "POP-1", "7.0", "7.0")]
        divergencias, sem_formula = divergencias_da_coluna_f(linhas, self._versoes())
        assert divergencias == []
        assert len(sem_formula) == 1 and "POP-1" in sem_formula[0]

    def test_linha_sem_codigo_e_ignorada(self):
        assert divergencias_da_coluna_f([(5, "  ", "x", "=x")], self._versoes()) == ([], [])

    def test_lista_vazia(self):
        assert divergencias_da_coluna_f(None, {}) == ([], [])
