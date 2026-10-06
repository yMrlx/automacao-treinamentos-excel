from treinamentos_its.relatorio import escrever_relatorio_no_rodape, montar_linhas_relatorio

# O rodapé leva o mês do CICLO (= do arquivo criado), regra do negócio de 06/10.
MESES = {"anterior": "Junho", "ano_ant_relativo": 2026, "atual": "Julho", "ano_atual": 2026}


# --- Planilha falsa mínima para exercitar escrever_relatorio_no_rodape --------
# Reproduz só o que a função usa da API do xlwings: used_range.last_cell.row,
# range(row, col) devolvendo uma célula, e range((r1,c1),(r2,c2)) devolvendo os
# valores de um intervalo de uma coluna.


class _FakeFont:
    def __init__(self):
        self.bold = False
        self.size = None


class _FakeCell:
    def __init__(self):
        self.value = None
        self.font = _FakeFont()


class _FakeIntervalo:
    def __init__(self, valores):
        self.value = valores


class _FakeLastCell:
    def __init__(self, row):
        self.row = row


class _FakeUsedRange:
    def __init__(self, sheet):
        self._sheet = sheet

    @property
    def last_cell(self):
        linhas = [linha for (linha, _coluna) in self._sheet.celulas] or [1]
        return _FakeLastCell(max(linhas))


class _FakeSheet:
    def __init__(self, name="Aba"):
        self.name = name
        self.celulas: dict[tuple[int, int], _FakeCell] = {}

    @property
    def used_range(self):
        return _FakeUsedRange(self)

    def range(self, *args):
        if len(args) == 2 and all(isinstance(a, int) for a in args):
            return self.celulas.setdefault((args[0], args[1]), _FakeCell())
        (linha_ini, col_ini), (linha_fim, _col_fim) = args
        valores = [
            self.celulas[(linha, col_ini)].value if (linha, col_ini) in self.celulas else None
            for linha in range(linha_ini, linha_fim + 1)
        ]
        return _FakeIntervalo(valores if len(valores) != 1 else valores[0])

    def coluna_a(self):
        return [
            cel.value
            for (linha, coluna), cel in sorted(self.celulas.items())
            if coluna == 1 and cel.value is not None
        ]


def test_escrever_relatorio_no_rodape_escreve_apos_os_dados():
    ws = _FakeSheet()
    ws.range(1, 1).value = "cabecalho"
    ws.range(2, 1).value = "dado"

    escrever_relatorio_no_rodape(ws, ["Referente Julho 2026", "Não houveram alterações nos POPs"])

    assert ws.coluna_a() == ["cabecalho", "dado", "Referente Julho 2026", "Não houveram alterações nos POPs"]


def test_escrever_relatorio_no_rodape_nao_reescreve_bloco_identico():
    ws = _FakeSheet()
    ws.range(1, 1).value = "dado"
    relatorio = ["Referente Julho 2026", "linha A", "linha B"]
    escrever_relatorio_no_rodape(ws, relatorio)
    coluna_apos_primeira = ws.coluna_a()

    escrever_relatorio_no_rodape(ws, relatorio)  # exatamente o mesmo conteúdo

    assert ws.coluna_a() == coluna_apos_primeira


def test_escrever_relatorio_no_rodape_reescreve_quando_o_conteudo_difere():
    # Mesmo título, conteúdo diferente (bloco herdado do mês anterior) -> escreve.
    ws = _FakeSheet()
    ws.range(1, 1).value = "dado"
    escrever_relatorio_no_rodape(ws, ["Referente Julho 2026", "versionado X"])

    escrever_relatorio_no_rodape(ws, ["Referente Julho 2026", "versionado Y"])

    assert ws.coluna_a() == ["dado", "Referente Julho 2026", "versionado X", "Referente Julho 2026", "versionado Y"]


def test_montar_linhas_relatorio_com_versionamento_e_atrasado():
    alteracoes = {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}
    atrasados = [{"nome": "FULANO", "codigo": "COD2"}]

    linhas = montar_linhas_relatorio(MESES, alteracoes, atrasados)

    assert linhas == [
        "Referente Julho 2026",
        "Versionado do treinamento COD1 de v1.0 para v2.0",
        "FULANO - COD2 ATRASADO",
    ]


def test_montar_linhas_relatorio_verificar_gera_linha_por_codigo():
    # A3 (2026-09-02): "VERIFICAR" não é mais alvo de versionamento - vira uma
    # linha própria no rodapé, pedindo pra um humano conferir.
    linhas = montar_linhas_relatorio(MESES, {}, [], verificar=["COD1", "COD2"])

    assert linhas == [
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
        "COD1 — POP obsoletado ou renomeado, para um humano verificar",
        "COD2 — POP obsoletado ou renomeado, para um humano verificar",
    ]


def test_montar_linhas_relatorio_verificar_junto_com_versionamento_e_atrasado():
    alteracoes = {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}
    atrasados = [{"nome": "FULANO", "codigo": "COD9"}]

    linhas = montar_linhas_relatorio(MESES, alteracoes, atrasados, verificar=["COD3", "COD2"])

    assert linhas == [
        "Referente Julho 2026",
        "Versionado do treinamento COD1 de v1.0 para v2.0",
        "COD2 — POP obsoletado ou renomeado, para um humano verificar",
        "COD3 — POP obsoletado ou renomeado, para um humano verificar",
        "FULANO - COD9 ATRASADO",
    ]


def test_montar_linhas_relatorio_sem_verificar_nao_adiciona_linha():
    linhas = montar_linhas_relatorio(MESES, {}, [])

    assert linhas == ["Referente Julho 2026", "Não houveram alterações nos POPs"]


def test_montar_linhas_relatorio_sem_alteracoes():
    linhas = montar_linhas_relatorio(MESES, {}, [])

    assert linhas == ["Referente Julho 2026", "Não houveram alterações nos POPs"]


def test_montar_linhas_relatorio_ordena_por_codigo():
    alteracoes = {
        "COD2": {"v_antiga": "1.0", "v_nova": "2.0"},
        "COD1": {"v_antiga": "3.0", "v_nova": "4.0"},
    }

    linhas = montar_linhas_relatorio(MESES, alteracoes, [])

    assert linhas[1:] == [
        "Versionado do treinamento COD1 de v3.0 para v4.0",
        "Versionado do treinamento COD2 de v1.0 para v2.0",
    ]


def test_montar_linhas_relatorio_rotulo_customizado():
    alteracoes = {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}

    linhas = montar_linhas_relatorio(MESES, alteracoes, [], rotulo_versionado="Versionado treinamento")

    assert linhas[1] == "Versionado treinamento COD1 de v1.0 para v2.0"
