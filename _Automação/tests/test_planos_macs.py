from datetime import datetime
from types import SimpleNamespace

from treinamentos_its.dominio import AvaliacaoMatriz
from treinamentos_its.planos_macs import (
    _atualizar_aba,
    _quitar_matriz_seletivamente,
    avisar_versionados_sem_bloco,
    codigos_versionados_sem_bloco,
    atualizar_planos_e_macs,
)


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
    def __init__(self, name="Cargo", rows=5, cols=5):
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


def _aba(status="OK", planejado=None):
    ws = _Sheet()
    ws.range(1, 1).value = "Treinamentos"
    ws.range(2, 3).value = "DOC-POP-1"
    ws.range(3, 3).value = "1.0"
    ws.range(5, 2).value = "Carla"
    ws.range(5, 3).value = planejado
    ws.range(5, 4).value = status
    return ws


def _avaliacao():
    return AvaliacaoMatriz(
        alteracoes={"DOC-POP-1": {"v_antiga": "1.0", "v_nova": "2.0"}},
        prazos={"DOC-POP-1": datetime(2026, 10, 20)},
        linhas_quitacao={"DOC-POP-1": 5},
    )


MESES = {
    "anterior": "Agosto",
    "ano_ant_relativo": 2026,
    "atual": "Setembro",
    "ano_atual": 2026,
}


def test_ok_vira_on_time_sem_escrever_na_linha_de_versao():
    ws = _aba("OK")
    codigos, marcacoes, _atrasados, desconhecidos = _atualizar_aba(
        ws, _avaliacao(), MESES, datetime(2026, 9, 25)
    )
    assert codigos == {"DOC-POP-1"}
    assert ws.range(3, 3).value == "1.0"
    assert ws.range(5, 4).value == "ON TIME"
    assert "20/10/2026" in ws.range(5, 3).value
    assert len(marcacoes) == 1
    assert desconhecidos == []


def test_elaborador_nao_e_cobrado():
    ws = _aba("OK", "ELABORADOR")
    _atualizar_aba(ws, _avaliacao(), MESES, datetime(2026, 9, 25))
    assert ws.range(5, 3).value == "ELABORADOR"
    assert ws.range(5, 4).value == "OK"


def test_hse_outros_e_global_nao_mudam():
    for status in ("HSE", "OUTROS", "GLOBAL"):
        ws = _aba(status)
        _atualizar_aba(ws, _avaliacao(), MESES, datetime(2026, 9, 25))
        assert ws.range(5, 4).value == status


def test_texto_desconhecido_e_preservado_e_reportado():
    ws = _aba("REVISAR")
    *_resto, desconhecidos = _atualizar_aba(
        ws, _avaliacao(), MESES, datetime(2026, 9, 25)
    )
    assert ws.range(5, 4).value == "REVISAR"
    assert "status='REVISAR'" in desconhecidos[0]


def test_versionado_sem_bloco_so_vira_warning_e_o_ciclo_pode_seguir(caplog):
    alteracoes = {"POP-2": {"v_antiga": "1", "v_nova": "2"}}
    assert codigos_versionados_sem_bloco(
        alteracoes, {"POP-1"}
    ) == ["POP-2"]
    assert avisar_versionados_sem_bloco(alteracoes, {"POP-1"}) == ["POP-2"]
    assert "versionou mas ninguém tem esse curso no Planos" in caplog.text
    assert "ciclo seguirá normalmente" in caplog.text


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


def test_quitacao_copia_f_para_e_na_linha_validada():
    matriz = _Sheet("Matriz - Atualização", rows=5, cols=6)
    matriz.range(5, 3).value = "DOC-POP-1"
    matriz.range(5, 5).value = "1.0"
    matriz.range(5, 6).value = 2.0
    matriz.range(5, 6).formula = "=PROCV()"
    wb = SimpleNamespace(sheets=_Sheets([matriz]))
    assert _quitar_matriz_seletivamente(wb, {"DOC-POP-1": 5}) == 1
    # Grava como TEXTO protegido (o Excel real guarda "2.0"; o dublê guarda o que
    # foi escrito). Antes ia o número 2.0 e o Excel mostrava "2" (05/10/2026).
    assert matriz.range(5, 5).value == "'2.0"


def test_ciclo_com_pop_sem_bloco_quita_e_termina_com_sucesso(
    monkeypatch, tmp_path, caplog
):
    from treinamentos_its import planos_macs as modulo

    cargo = _Sheet("Cargo", rows=1, cols=1)
    matriz = _Sheet("Matriz - Atualização", rows=5, cols=6)
    matriz.range(5, 3).value = "DOC-POP-ORFAO"
    matriz.range(5, 5).value = "1.0"
    matriz.range(5, 6).value = 2.0
    matriz.range(5, 6).formula = "=PROCV()"
    wb = _Workbook([cargo, matriz])

    class _App:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

    arquivo = tmp_path / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("fake")
    avaliacao = AvaliacaoMatriz(
        alteracoes={
            "DOC-POP-ORFAO": {"v_antiga": "1.0", "v_nova": "2.0"}
        },
        prazos={"DOC-POP-ORFAO": datetime(2026, 10, 20)},
        linhas_quitacao={"DOC-POP-ORFAO": 5},
    )
    monkeypatch.setattr(modulo, "ABAS_DE_CARGO", ("Cargo",))
    monkeypatch.setattr(modulo, "app_excel", lambda: _App())
    monkeypatch.setattr(modulo, "abrir_livro", lambda *_args, **_kwargs: wb)

    assert atualizar_planos_e_macs(
        avaliacao, MESES, arquivo, hoje=datetime(2026, 9, 25)
    )
    assert wb.salvou
    assert matriz.range(5, 5).value == "'2.0"  # quitação grava texto protegido
    assert "versionou mas ninguém tem esse curso no Planos" in caplog.text


# =============================================================================
# Cobertura recriada depois do a5dad36 (adaptada ao código SEM Procedimentos).
# Tudo com a planilha de mentira de `dubles_excel` - nenhum teste abre Excel.
# =============================================================================

import logging

import pytest

from dubles_excel import (
    FORMULA_F,
    AbaFalsa,
    ExcelFalso,
    LivroFalso,
    aba_de_cargo,
    aba_matriz,
    abas_de_cargo_vazias,
    celulas_da_pessoa,
)
from treinamentos_its import planos_macs as pm
from treinamentos_its.config import ABAS_DE_CARGO
from treinamentos_its.dominio import avaliar_matriz

CODIGO = "DOC-POP-0000001"
OUTRO = "DOC-POP-0000002"
ANALISTA = ABAS_DE_CARGO[0]
GERENTE = ABAS_DE_CARGO[3]
HOJE = datetime(2026, 9, 25)
PRAZO_FUTURO = datetime(2026, 10, 20)
PRAZO_VENCIDO = datetime(2026, 9, 1)


def _versionando(*codigos, prazo=PRAZO_FUTURO, verificar=()):
    return AvaliacaoMatriz(
        alteracoes={codigo: {"v_antiga": "1.0", "v_nova": "2.0"} for codigo in codigos},
        prazos={codigo: prazo for codigo in codigos},
        linhas_quitacao={codigo: 5 + indice for indice, codigo in enumerate(codigos)},
        verificar=list(verificar),
    )


def _bloco_puro(linha_cod: int, codigos: dict[int, str]) -> dict:
    """Bloco no formato de `_ler_blocos` (as funções testadas só leem 'codigos')."""
    return {"linha_cod": linha_cod, "linha_func": linha_cod + 3, "codigos": codigos}


def _livro_do_ciclo(*abas_cargo: AbaFalsa, matriz: AbaFalsa | None = None) -> LivroFalso:
    nomes = tuple(aba.name for aba in abas_cargo)
    matriz = matriz or aba_matriz([(CODIGO, "1.0", 2.0, FORMULA_F)])
    return LivroFalso(*abas_cargo, *abas_de_cargo_vazias(exceto=nomes), matriz)


def _arquivo(tmp_path):
    arquivo = tmp_path / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_text("marcador de existência", encoding="utf-8")
    return arquivo


def _texto(planejado) -> str:
    return str(planejado or "").replace("'", "")


# --- _e_cabecalho_treinamentos (B4) -------------------------------------------


class TestCabecalhoTreinamentos:
    def test_palavra_exata(self):
        assert pm._e_cabecalho_treinamentos("Treinamentos") is True

    def test_tolera_caixa(self):
        assert pm._e_cabecalho_treinamentos("TREINAMENTOS") is True
        assert pm._e_cabecalho_treinamentos("treinamentos") is True

    def test_tolera_espaco_nas_bordas(self):
        assert pm._e_cabecalho_treinamentos("  Treinamentos ") is True
        assert pm._e_cabecalho_treinamentos("treinamentos ") is True

    def test_tolera_acento(self):
        assert pm._e_cabecalho_treinamentos("Treinâmentos") is True

    def test_combina_caixa_espaco_e_acento(self):
        assert pm._e_cabecalho_treinamentos("  TREINÂMENTOS  ") is True

    def test_falso_para_vazio_none_e_zero(self):
        assert pm._e_cabecalho_treinamentos(None) is False
        assert pm._e_cabecalho_treinamentos("") is False
        assert pm._e_cabecalho_treinamentos(0) is False

    def test_falso_para_outro_texto(self):
        assert pm._e_cabecalho_treinamentos("Treinamento") is False  # singular
        assert pm._e_cabecalho_treinamentos("Treinamentos obrigatórios") is False
        assert pm._e_cabecalho_treinamentos("Trainings") is False
        assert pm._e_cabecalho_treinamentos("ID") is False

    def test_valor_numerico_nao_quebra(self):
        assert pm._e_cabecalho_treinamentos(123) is False


# --- dados_tem_conteudo (R2: aba vazia x aba com dados) -----------------------


class TestDadosTemConteudo:
    def test_aba_vazia(self):
        # Aba vazia é legítima (categoria sem ninguém) -> não vira trava.
        assert pm.dados_tem_conteudo(None) is False
        assert pm.dados_tem_conteudo([]) is False
        assert pm.dados_tem_conteudo([[None, None], [None, ""]]) is False

    def test_qualquer_celula_preenchida(self):
        assert pm.dados_tem_conteudo([[None, None], [None, "Fulano"]]) is True
        assert pm.dados_tem_conteudo([[0, None]]) is True

    def test_linha_unica_nao_aninhada(self):
        # O xlwings devolve lista simples quando a faixa tem uma linha só.
        assert pm.dados_tem_conteudo(["Treinamentos", None]) is True
        assert pm.dados_tem_conteudo([None, ""]) is False


# --- codigos_dos_blocos / filtrar_verificar_da_aba (A3) -----------------------


class TestFiltrarVerificarDaAba:
    def test_codigos_dos_blocos_junta_todos_os_blocos(self):
        blocos = [
            _bloco_puro(5, {3: "DOC-POP-001", 5: "DOC-POP-002"}),
            _bloco_puro(20, {3: "DOC-CRG-010"}),
        ]
        assert pm.codigos_dos_blocos(blocos) == {
            "DOC-POP-001",
            "DOC-POP-002",
            "DOC-CRG-010",
        }

    def test_codigos_dos_blocos_sem_blocos(self):
        assert pm.codigos_dos_blocos([]) == set()
        assert pm.codigos_dos_blocos(None) == set()

    def test_mantem_so_quem_aparece_na_aba(self):
        blocos = [_bloco_puro(5, {3: "DOC-POP-001", 5: "DOC-POP-002"})]
        assert pm.filtrar_verificar_da_aba(
            ["DOC-POP-002", "DOC-POP-099", "DOC-POP-001"], blocos
        ) == ["DOC-POP-002", "DOC-POP-001"]

    def test_nao_depende_de_versionamento(self):
        # VERIFICAR nunca versiona ninguém: filtrar por `alteracoes` zeraria a
        # lista sempre. Sem versionamento nenhum, o código tem que passar.
        assert pm.filtrar_verificar_da_aba(
            ["DOC-POP-001"], [_bloco_puro(5, {3: "DOC-POP-001"})]
        ) == ["DOC-POP-001"]

    def test_sem_codigos_ou_sem_blocos(self):
        bloco = [_bloco_puro(5, {3: "DOC-POP-001"})]
        assert pm.filtrar_verificar_da_aba([], bloco) == []
        assert pm.filtrar_verificar_da_aba(None, bloco) == []
        assert pm.filtrar_verificar_da_aba(["DOC-POP-001"], []) == []

    def test_ignora_espaco_vazio_e_duplicata(self):
        blocos = [_bloco_puro(5, {3: "DOC-POP-001"}), _bloco_puro(20, {3: "DOC-POP-001"})]
        assert pm.filtrar_verificar_da_aba(
            ["  DOC-POP-001 ", "DOC-POP-001", "", None], blocos
        ) == ["DOC-POP-001"]

    def test_preserva_a_ordem_de_entrada(self):
        blocos = [_bloco_puro(5, {3: "DOC-POP-003", 5: "DOC-POP-001", 7: "DOC-POP-002"})]
        assert pm.filtrar_verificar_da_aba(["DOC-POP-002", "DOC-POP-003"], blocos) == [
            "DOC-POP-002",
            "DOC-POP-003",
        ]


# --- travas das abas de cargo (aba faltando / dados sem bloco) ----------------


def _aba_sem_bloco(nome: str) -> AbaFalsa:
    # Cabeçalho no singular: o bloco não é reconhecido, mas a aba TEM gente.
    aba = AbaFalsa(nome)
    aba.celula(1, 1).value = "Treinamento"
    aba.celula(2, 3).value = CODIGO
    aba.celula(5, 2).value = "Carla"
    aba.celula(5, 4).value = "OK"
    return aba


class TestTravasDasAbasDeCargo:
    def test_todas_presentes_com_bloco_ou_vazias_nao_trava(self):
        livro = _livro_do_ciclo(aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)))
        assert pm.travas_das_abas_de_cargo(livro) == []

    def test_desliga_o_autofiltro_em_memoria_antes_de_ler(self):
        aba = aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))
        aba.api.AutoFilterMode = True
        livro = _livro_do_ciclo(aba)

        assert pm.travas_das_abas_de_cargo(livro) == []
        assert aba.api.AutoFilterMode is False

    def test_aba_de_cargo_faltando_vira_trava(self):
        livro = LivroFalso(*abas_de_cargo_vazias(exceto=(GERENTE,)))
        travas = pm.travas_das_abas_de_cargo(livro)
        assert len(travas) == 1
        assert "não encontrada" in travas[0] and GERENTE in travas[0]

    def test_aba_com_dados_e_sem_bloco_vira_trava(self):
        livro = _livro_do_ciclo(_aba_sem_bloco(ANALISTA))
        travas = pm.travas_das_abas_de_cargo(livro)
        assert len(travas) == 1
        assert ANALISTA in travas[0]
        assert "nenhum bloco 'Treinamentos'" in travas[0]

    def test_todas_as_travas_sao_listadas_juntas(self):
        livro = LivroFalso(
            _aba_sem_bloco(ANALISTA),
            *abas_de_cargo_vazias(exceto=(ANALISTA, GERENTE)),
        )
        travas = pm.travas_das_abas_de_cargo(livro)
        assert len(travas) == 2
        assert any(GERENTE in trava for trava in travas)
        assert any(ANALISTA in trava for trava in travas)

    def test_atualizar_aba_com_dados_sem_bloco_levanta(self):
        # Defesa em profundidade: mesmo chamada direto, a aba quebrada não passa.
        with pytest.raises(RuntimeError, match="nenhum bloco 'Treinamentos'"):
            pm._atualizar_aba(_aba_sem_bloco(ANALISTA), _versionando(CODIGO), MESES, HOJE)

    @pytest.mark.parametrize("quebra", ["aba faltando", "aba sem bloco"])
    def test_ciclo_para_antes_da_primeira_escrita(
        self, monkeypatch, tmp_path, caplog, quebra
    ):
        # A aba boa é a PRIMEIRA do ABAS_DE_CARGO e a quebrada vem depois: sem
        # a checagem prévia, a boa já teria sido escrita quando o erro aparece.
        assert ABAS_DE_CARGO.index(ANALISTA) < ABAS_DE_CARGO.index(GERENTE)
        boa = aba_de_cargo(ANALISTA, [("Caio", {})], (CODIGO,))
        vazias = abas_de_cargo_vazias(exceto=(GERENTE, ANALISTA))
        if quebra == "aba faltando":
            abas = [boa, *vazias]
        else:
            abas = [boa, _aba_sem_bloco(GERENTE), *vazias]
        livro = LivroFalso(*abas, aba_matriz([(CODIGO, "1.0", 2.0, FORMULA_F)]))
        antes = boa.estado()
        ExcelFalso(livro).instalar(monkeypatch, pm)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert pm.atualizar_planos_e_macs(
                _versionando(CODIGO), MESES, _arquivo(tmp_path), hoje=HOJE
            ) is False

        assert "O ciclo vai PARAR por isto" in caplog.text
        assert boa.estado() == antes  # nenhuma aba foi tocada
        assert livro.sheets["Matriz - Atualização"].celula(5, 5).value == "1.0"
        assert livro.salvamentos == 0
        assert livro.fechamentos == 1


# --- recusar antes de abrir o Excel --------------------------------------------


def _excel_proibido(*_args, **_kwargs):
    raise AssertionError("não era pra abrir o Excel")


class TestRecusaSemAbrirExcel:
    def test_matriz_com_bloqueio(self, monkeypatch, tmp_path, caplog):
        monkeypatch.setattr(pm, "app_excel", _excel_proibido)
        avaliacao = AvaliacaoMatriz(bloqueios=["regressão de versão: QU (2.0 -> 1.0)"])
        assert pm.atualizar_planos_e_macs(avaliacao, MESES, _arquivo(tmp_path)) is False
        assert "A Matriz tem bloqueios" in caplog.text

    def test_arquivo_inexistente(self, monkeypatch, tmp_path):
        monkeypatch.setattr(pm, "app_excel", _excel_proibido)
        assert pm.atualizar_planos_e_macs(
            _versionando(CODIGO), MESES, tmp_path / "nao-existe.xlsx"
        ) is False


# --- o que fica escrito em cada célula -----------------------------------------


class TestMarcacaoNaPlanilha:
    def _rodar(self, planejado, status, avaliacao=None, nome="Carla"):
        ws = aba_de_cargo(ANALISTA, [(nome, {CODIGO: (planejado, status)})], (CODIGO,))
        resultado = pm._atualizar_aba(ws, avaliacao or _versionando(CODIGO), MESES, HOJE)
        return ws, resultado

    def test_versionado_atrasado_continua_com_a_data_antiga(self):
        ws, (_c, marcacoes, atrasados, _d) = self._rodar("'01/08/2026", "ATRASADO")
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert planejado.value == "'01/08/2026"  # data antiga intocada
        assert executado.value == "ATRASADO"
        assert executado.font.color == (255, 0, 0)
        assert [m.status for m in marcacoes] == ["ATRASADO"]
        assert marcacoes[0].planejado == "'01/08/2026"
        assert atrasados == [{"nome": "CARLA", "codigo": CODIGO}]

    def test_versionado_on_time_vencido_vira_atrasado_mantendo_a_data(self):
        ws, (_c, marcacoes, atrasados, _d) = self._rodar(datetime(2026, 9, 10), "ON TIME")
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert planejado.value == datetime(2026, 9, 10)
        assert executado.value == "ATRASADO"
        assert marcacoes[0].planejado == datetime(2026, 9, 10)
        assert len(atrasados) == 1

    def test_versionado_on_time_no_prazo_ganha_o_prazo_novo(self):
        ws, _resultado = self._rodar("'30/09/2026", "ON TIME")
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert _texto(planejado.value) == "20/10/2026"
        assert executado.value == "ON TIME"

    def test_versionado_ok_com_prazo_vencido_nasce_atrasado_com_prazo_novo(self):
        ws, (_c, _m, atrasados, _d) = self._rodar(
            None, "OK", _versionando(CODIGO, prazo=PRAZO_VENCIDO)
        )
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert _texto(planejado.value) == "01/09/2026"
        assert executado.value == "ATRASADO"
        assert len(atrasados) == 1

    def test_nao_versionado_on_time_vencido_vira_atrasado_mantendo_a_data(self):
        ws, (_c, marcacoes, atrasados, _d) = self._rodar(
            "'10/09/2026", "ON TIME", AvaliacaoMatriz()
        )
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert planejado.value == "'10/09/2026"
        assert executado.value == "ATRASADO"
        assert [m.status for m in marcacoes] == ["ATRASADO"]
        assert atrasados == [{"nome": "CARLA", "codigo": CODIGO}]

    def test_nao_versionado_on_time_no_prazo_nao_muda(self):
        ws, (_c, marcacoes, atrasados, _d) = self._rodar(
            "'30/10/2026", "ON TIME", AvaliacaoMatriz()
        )
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert (planejado.value, executado.value) == ("'30/10/2026", "ON TIME")
        assert marcacoes == [] and atrasados == []

    def test_nao_versionado_atrasado_nao_e_remarcado(self):
        ws, (_c, marcacoes, _a, _d) = self._rodar("'01/08/2026", "ATRASADO", AvaliacaoMatriz())
        _planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert executado.value == "ATRASADO"
        assert executado.font.color is None  # nem a cor foi repintada
        assert marcacoes == []

    @pytest.mark.parametrize("vazio", [None, ""])
    @pytest.mark.parametrize("texto", ["ELABORADOR", "APROVADORA"])
    def test_status_vazio_com_elaborador_ou_aprovadora_nao_mexe(self, vazio, texto):
        # Correção D do 510a997: vazio = OK, e OK + ELABORADOR/APROVADORA não mexe.
        ws, (_c, marcacoes, _a, desconhecidos) = self._rodar(texto, vazio)
        planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert planejado.value == texto
        assert executado.value == vazio
        assert marcacoes == [] and desconhecidos == []

    def test_status_vazio_sem_planejado_vira_on_time(self):
        ws, _resultado = self._rodar(None, None)
        _planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert executado.value == "ON TIME"

    @pytest.mark.parametrize("status", ["NA", "N/A"])
    def test_na_nao_muda(self, status):
        ws, (_c, marcacoes, _a, desconhecidos) = self._rodar(None, status)
        _planejado, executado = celulas_da_pessoa(ws, "Carla")
        assert executado.value == status
        assert marcacoes == [] and desconhecidos == []

    def test_linha_versao_do_bloco_nunca_e_escrita(self):
        ws, _resultado = self._rodar(None, "OK")
        celula = ws.celula(3, 3)
        assert celula.value == 1.0
        assert celula.formula.startswith("=VLOOKUP")


# --- rodapé -----------------------------------------------------------------------


class TestRodape:
    def test_atrasado_no_rodape_em_caixa_alta_sem_acento(self):
        ws = aba_de_cargo(
            ANALISTA,
            [("João da Conceição", {CODIGO: ("'01/08/2026", "ATRASADO")})],
            (CODIGO,),
        )
        pm._atualizar_aba(ws, _versionando(CODIGO), MESES, HOJE)
        coluna_a = ws.coluna_a()
        assert f"JOAO DA CONCEICAO - {CODIGO} ATRASADO" in coluna_a
        assert not any("João" in str(texto) for texto in coluna_a)

    def test_rodape_tem_titulo_versionado_e_atrasado_na_ordem(self):
        ws = aba_de_cargo(
            ANALISTA,
            [("Carla", {CODIGO: ("'01/08/2026", "ATRASADO")}), ("Caio", {})],
            (CODIGO,),
        )
        pm._atualizar_aba(ws, _versionando(CODIGO), MESES, HOJE)
        coluna_a = ws.coluna_a()
        inicio = coluna_a.index("Referente Setembro 2026")
        assert coluna_a[inicio:] == [
            "Referente Setembro 2026",
            f"Versionado treinamento {CODIGO} de v1.0 para v2.0",
            f"CARLA - {CODIGO} ATRASADO",
        ]
        titulo = ws.celula(inicio + 1, 1)
        assert titulo.font.bold is True and titulo.font.size == 12

    def test_versionado_de_outra_aba_nao_entra_no_rodape_desta(self):
        ws = aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))
        pm._atualizar_aba(ws, _versionando(CODIGO, OUTRO), MESES, HOJE)
        texto = "\n".join(str(t) for t in ws.coluna_a() if t)
        assert CODIGO in texto and OUTRO not in texto

    def test_verificar_so_vai_pro_rodape_da_aba_onde_o_codigo_aparece(self):
        ws = aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))
        pm._atualizar_aba(
            ws, AvaliacaoMatriz(verificar=[CODIGO, OUTRO]), MESES, HOJE
        )
        coluna_a = ws.coluna_a()
        assert f"{CODIGO} — POP obsoletado ou renomeado, para um humano verificar" in coluna_a
        assert not any(OUTRO in str(texto) for texto in coluna_a)

    @pytest.mark.parametrize(
        "planejado, status, prazo",
        [
            (None, "OK", PRAZO_FUTURO),
            (None, "OK", PRAZO_VENCIDO),
            ("'01/08/2026", "ATRASADO", PRAZO_FUTURO),
        ],
        ids=["on time", "nasce atrasado", "ja atrasado"],
    )
    def test_rodar_duas_vezes_nao_repete_o_bloco(self, planejado, status, prazo):
        ws = aba_de_cargo(ANALISTA, [("Carla", {CODIGO: (planejado, status)})], (CODIGO,))
        avaliacao = _versionando(CODIGO, prazo=prazo, verificar=[CODIGO])
        pm._atualizar_aba(ws, avaliacao, MESES, HOJE)
        depois_da_primeira = ws.coluna_a()

        pm._atualizar_aba(ws, avaliacao, MESES, HOJE)

        assert ws.coluna_a() == depois_da_primeira
        assert ws.coluna_a().count("Referente Setembro 2026") == 1

    def test_rodape_herdado_com_o_mesmo_titulo_e_preservado(self):
        # Bloco antigo com o mesmo título e conteúdo diferente é auditoria:
        # fica, e o bloco deste ciclo vai embaixo.
        herdado = ("Referente Setembro 2026", "Não houveram alterações nos POPs")
        ws = aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,), rodape=herdado)
        pm._atualizar_aba(ws, _versionando(CODIGO), MESES, HOJE)
        coluna_a = ws.coluna_a()
        assert coluna_a.count("Referente Setembro 2026") == 2
        assert "Não houveram alterações nos POPs" in coluna_a


# --- convergência: o mesmo ciclo 3x na mesma planilha ------------------------------


CENARIOS_DA_ESPEC = [
    # (pessoa, planejado, status) - tabela §2.2 da espec
    ("Ok Futuro", None, "OK"),
    ("Ok Elaborador", "ELABORADOR", "OK"),
    ("Ok Aprovadora", "APROVADORA", "OK"),
    ("On Time Vencido", "'10/09/2026", "ON TIME"),
    ("On Time No Prazo", "'30/09/2026", "ON TIME"),
    ("Atrasado", "'01/08/2026", "ATRASADO"),
    ("Na", None, "NA"),
    ("N Barra A", None, "N/A"),
    ("Hse", None, "HSE"),
    ("Outros", None, "OUTROS"),
    ("Global", None, "GLOBAL"),
    ("Vazio", None, None),
    ("Vazio Elaborador", "ELABORADOR", None),
    ("Texto Estranho", None, "REVISAR"),
]
NAO_MEXE = {
    "Ok Elaborador", "Ok Aprovadora", "Na", "N Barra A", "Hse", "Outros", "Global",
    "Vazio Elaborador", "Texto Estranho",
}


class TestConvergenciaNaPlanilha:
    @pytest.mark.parametrize(
        "prazo", [PRAZO_FUTURO, PRAZO_VENCIDO], ids=["prazo futuro", "prazo vencido"]
    )
    def test_mesmo_ciclo_tres_vezes_da_o_mesmo_resultado(self, prazo):
        ws = aba_de_cargo(
            ANALISTA,
            [(nome, {CODIGO: (planejado, status)}) for nome, planejado, status in CENARIOS_DA_ESPEC],
            (CODIGO,),
        )
        avaliacao = _versionando(CODIGO, prazo=prazo)
        estados = []
        for _vez in range(3):
            pm._atualizar_aba(ws, avaliacao, MESES, HOJE)
            estados.append(ws.estado())

        assert estados[0] == estados[1] == estados[2]

        def status(nome):
            return celulas_da_pessoa(ws, nome)[1].value

        novo = "ON TIME" if prazo is PRAZO_FUTURO else "ATRASADO"
        assert status("Ok Futuro") == novo
        assert status("Vazio") == novo
        assert status("On Time No Prazo") == novo
        assert status("On Time Vencido") == "ATRASADO"
        assert status("Atrasado") == "ATRASADO"
        for nome, _planejado, original in CENARIOS_DA_ESPEC:
            if nome in NAO_MEXE:
                assert status(nome) == original, nome


# --- fluxo inteiro com a planilha de mentira -------------------------------------


class TestFluxoAtualizarPlanosEMacs:
    def test_status_desconhecido_e_a_ultima_coisa_logada(self, monkeypatch, tmp_path, caplog):
        # Correção E do 510a997: a lista vai no FIM, completa (das duas abas).
        analista = aba_de_cargo(
            ANALISTA, [("Carla", {CODIGO: (None, "REVISAR")}), ("Caio", {})], (CODIGO,)
        )
        gerente = aba_de_cargo(GERENTE, [("Dora", {CODIGO: (None, "???")})], (CODIGO,))
        livro = _livro_do_ciclo(analista, gerente)
        ExcelFalso(livro).instalar(monkeypatch, pm)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert pm.atualizar_planos_e_macs(
                _versionando(CODIGO), MESES, _arquivo(tmp_path), hoje=HOJE
            ) is True

        mensagens = [registro.getMessage() for registro in caplog.records]
        sucesso = mensagens.index("Planos e Macs atualizado com sucesso.")
        finais = caplog.records[sucesso + 1:]
        assert [registro.levelno for registro in finais] == [logging.WARNING] * 3
        assert "2 status desconhecido(s)" in finais[0].getMessage()
        itens = finais[1].getMessage() + finais[2].getMessage()
        assert "'REVISAR'" in itens and "'???'" in itens
        assert ANALISTA in itens and GERENTE in itens
        # e eles não foram mexidos
        assert celulas_da_pessoa(analista, "Carla")[1].value == "REVISAR"
        assert celulas_da_pessoa(gerente, "Dora")[1].value == "???"
        assert celulas_da_pessoa(analista, "Caio")[1].value == "ON TIME"

    def test_verificar_fora_de_todos_os_blocos_avisa_com_o_arquivo(
        self, monkeypatch, tmp_path, caplog
    ):
        # Correção K do 510a997.
        livro = _livro_do_ciclo(aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)))
        ExcelFalso(livro).instalar(monkeypatch, pm)
        arquivo = _arquivo(tmp_path)
        avaliacao = _versionando(CODIGO, verificar=["DOC-POP-9999999"])

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert pm.atualizar_planos_e_macs(avaliacao, MESES, arquivo, hoje=HOJE)

        avisos = [r.getMessage() for r in caplog.records if r.levelno == logging.WARNING]
        assert any(
            "DOC-POP-9999999" in aviso and "'VERIFICAR'" in aviso and arquivo.name in aviso
            for aviso in avisos
        )

    def test_ciclo_completo_desliga_autosave_salva_uma_vez_e_reabre_so_leitura(
        self, monkeypatch, tmp_path
    ):
        analista = aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,))
        livro = _livro_do_ciclo(analista)
        excel = ExcelFalso(livro).instalar(monkeypatch, pm)
        arquivo = _arquivo(tmp_path)

        assert pm.atualizar_planos_e_macs(_versionando(CODIGO), MESES, arquivo, hoje=HOJE)

        assert livro.api.AutoSaveOn is False
        assert livro.salvamentos == 1
        assert excel.aberturas == [(arquivo, False), (arquivo, True)]
        assert livro.sheets["Matriz - Atualização"].celula(5, 5).value == "'2.0"  # texto protegido
        assert celulas_da_pessoa(analista, "Carla")[1].value == "ON TIME"

    def test_quitacao_que_falha_no_meio_nao_salva(self, monkeypatch, tmp_path, caplog):
        # A linha da Matriz mudou entre a validação e a quitação.
        matriz = aba_matriz([("DOC-POP-OUTRO", "1.0", 2.0, FORMULA_F)])
        livro = _livro_do_ciclo(aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)), matriz=matriz)
        ExcelFalso(livro).instalar(monkeypatch, pm)

        assert pm.atualizar_planos_e_macs(
            _versionando(CODIGO), MESES, _arquivo(tmp_path), hoje=HOJE
        ) is False
        assert livro.salvamentos == 0
        assert "mudou antes de registrar a versão nova" in caplog.text


class TestAvisarVerificarSemBloco:
    def test_lista_so_os_que_nao_aparecem(self, caplog):
        orfaos = pm.avisar_verificar_sem_bloco(
            ["POP-9", "POP-1", " POP-9 ", "", None], {"POP-1"}, "Planos.xlsx"
        )
        assert orfaos == ["POP-9"]
        assert "POP-9" in caplog.text and "Planos.xlsx" in caplog.text

    def test_sem_orfao_nao_avisa(self, caplog):
        assert pm.avisar_verificar_sem_bloco(["POP-1"], {"POP-1"}, "Planos.xlsx") == []
        assert pm.avisar_verificar_sem_bloco(None, set(), "Planos.xlsx") == []
        assert caplog.records == []


# --- quitação seletiva E = F ---------------------------------------------------------


def _livro_com_matriz(linhas) -> LivroFalso:
    return LivroFalso(aba_matriz(linhas))


class TestQuitacaoSeletiva:
    def test_avaliacao_so_manda_quitar_f_numerica_valida(self):
        linhas = [
            ("POP-VERIFICAR", "1.0", "VERIFICAR"),
            ("POP-VAZIA", "1.0", None),
            ("POP-TEXTO", "1.0", "N/A"),
            ("POP-NEGATIVO", "1.0", -2146826246),  # código de erro COM do Excel
            ("POP-OK", "1.0", 2.0),
            ("POP-IGUAL", "3.0", 3.0),
        ]
        resultado = avaliar_matriz(linhas, formulas_f=[FORMULA_F] * len(linhas))
        assert resultado.linhas_quitacao == {"POP-OK": 9}

    @pytest.mark.parametrize(
        "f_na_hora",
        ["VERIFICAR", None, "N/A", -2146826246],
        ids=["verificar", "vazia", "texto", "negativo"],
    )
    def test_f_que_deixou_de_ser_numero_nao_quita_e_para(self, f_na_hora):
        # A linha foi validada como versionamento, mas na hora de quitar a F
        # já não é versão: E NÃO pode receber isso.
        livro = _livro_com_matriz([("POP-1", "1.0", f_na_hora, FORMULA_F)])
        with pytest.raises(RuntimeError, match="mudou antes de registrar a versão nova"):
            pm._quitar_matriz_seletivamente(livro, {"POP-1": 5})
        assert livro.sheets["Matriz - Atualização"].celula(5, 5).value == "1.0"

    def test_f_digitada_a_mao_nao_quita(self):
        livro = _livro_com_matriz([("POP-1", "1.0", 2.0, None)])
        assert livro.sheets["Matriz - Atualização"].celula(5, 6).formula == "2.0"
        with pytest.raises(RuntimeError, match="mudou antes de registrar a versão nova"):
            pm._quitar_matriz_seletivamente(livro, {"POP-1": 5})
        assert livro.sheets["Matriz - Atualização"].celula(5, 5).value == "1.0"

    def test_e_que_virou_texto_nao_quita(self):
        livro = _livro_com_matriz([("POP-1", "ver depois", 2.0, FORMULA_F)])
        with pytest.raises(RuntimeError, match="mudou antes de registrar a versão nova"):
            pm._quitar_matriz_seletivamente(livro, {"POP-1": 5})

    def test_aba_matriz_ausente(self):
        with pytest.raises(RuntimeError, match="não encontrada"):
            pm._quitar_matriz_seletivamente(LivroFalso(AbaFalsa("Outra")), {"POP-1": 5})

    def test_linha_com_outro_codigo_da_erro_com_linha_e_codigo(self):
        livro = _livro_com_matriz([("POP-2", "1.0", 2.0, FORMULA_F)])
        with pytest.raises(RuntimeError, match=r"linha 5 de POP-1 mudou"):
            pm._quitar_matriz_seletivamente(livro, {"POP-1": 5})
        assert livro.sheets["Matriz - Atualização"].celula(5, 5).value == "1.0"

    def test_quita_so_as_linhas_pedidas_e_cola_o_valor_bruto(self, caplog):
        livro = _livro_com_matriz(
            [
                ("POP-1", "1.0", 2.0, FORMULA_F),
                ("POP-2", "3.0", " 10.5 ", FORMULA_F),
                ("POP-3", "1.0", 5.0, FORMULA_F),  # não pedida
            ]
        )
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert pm._quitar_matriz_seletivamente(livro, {"POP-1": 5, "POP-2": 6}) == 2
        matriz = livro.sheets["Matriz - Atualização"]
        assert matriz.celula(5, 5).value == "'2.0"  # texto protegido (ver acima)
        assert matriz.celula(6, 5).value == "' 10.5 "  # texto como veio, protegido do Excel
        assert matriz.celula(7, 5).value == "1.0"
        assert matriz.celula(5, 6).formula == FORMULA_F  # F continua fórmula
        assert "Versões novas registradas na Matriz" in caplog.text and "2 linha(s)" in caplog.text

    def test_nada_a_quitar(self):
        assert pm._quitar_matriz_seletivamente(_livro_com_matriz([]), {}) == 0


# --- conferência depois de reabrir (planos_macs._reabrir_e_conferir) ---------------


class TestReabrirEConferir:
    def _cenario(self, monkeypatch, *, e=2.0, f=2.0, status="ON TIME", planejado="'20/10/2026"):
        ws = aba_de_cargo(ANALISTA, [("Carla", {CODIGO: (planejado, status)})], (CODIGO,))
        livro = LivroFalso(ws, aba_matriz([(CODIGO, e, f, FORMULA_F)]))
        excel = ExcelFalso(livro).instalar(monkeypatch, pm)
        marcacao = pm.MarcacaoAplicada(ANALISTA, 5, 3, PRAZO_FUTURO, "ON TIME")
        return livro, excel, _versionando(CODIGO), [marcacao]

    def test_tudo_salvo_confere(self, monkeypatch, tmp_path, caplog):
        livro, excel, avaliacao, marcacoes = self._cenario(monkeypatch)
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert pm._reabrir_e_conferir(tmp_path / "x.xlsx", avaliacao, marcacoes) is True
        assert excel.aberturas == [(tmp_path / "x.xlsx", True)]
        assert livro.fechamentos == 1
        assert "Planos e Macs conferido após reabrir" in caplog.text

    def test_quitacao_que_nao_ficou_salva(self, monkeypatch, tmp_path, caplog):
        livro, _excel, avaliacao, marcacoes = self._cenario(monkeypatch, e=1.0)
        assert pm._reabrir_e_conferir(tmp_path / "x.xlsx", avaliacao, marcacoes) is False
        assert "não permaneceu salva na Matriz" in caplog.text
        assert livro.fechamentos == 1

    def test_status_da_amostra_que_nao_ficou_salvo(self, monkeypatch, tmp_path):
        _l, _e, avaliacao, marcacoes = self._cenario(monkeypatch, status="OK")
        assert pm._reabrir_e_conferir(tmp_path / "x.xlsx", avaliacao, marcacoes) is False

    def test_prazo_da_amostra_que_nao_ficou_salvo(self, monkeypatch, tmp_path):
        _l, _e, avaliacao, marcacoes = self._cenario(monkeypatch, planejado="'21/10/2026")
        assert pm._reabrir_e_conferir(tmp_path / "x.xlsx", avaliacao, marcacoes) is False

    @pytest.mark.parametrize(
        "lido", ["'20/10/2026", "20/10/2026", datetime(2026, 10, 20, 0, 0)]
    )
    def test_prazo_em_texto_ou_data_confere(self, monkeypatch, tmp_path, lido):
        _l, _e, avaliacao, marcacoes = self._cenario(monkeypatch, planejado=lido)
        assert pm._reabrir_e_conferir(tmp_path / "x.xlsx", avaliacao, marcacoes) is True

    def test_no_fluxo_reabertura_divergente_faz_o_ciclo_falhar(self, monkeypatch, tmp_path):
        # O que foi salvo (livro reaberto) não é o que foi escrito: o disco
        # perdeu a quitação. O ciclo tem que falhar.
        escrito = _livro_do_ciclo(aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)))
        no_disco = _livro_do_ciclo(aba_de_cargo(ANALISTA, [("Carla", {})], (CODIGO,)))

        class _ExcelQuePerde(ExcelFalso):
            def abrir_livro(self, _app, caminho, *, read_only=False):
                self.aberturas.append((caminho, read_only))
                return no_disco if read_only else escrito

        _ExcelQuePerde(escrito).instalar(monkeypatch, pm)
        assert pm.atualizar_planos_e_macs(
            _versionando(CODIGO), MESES, _arquivo(tmp_path), hoje=HOJE
        ) is False
