"""Dublês do xlwings para os testes que NÃO abrem Excel.

Não é arquivo de teste (o pytest só coleta ``test_*.py``): é a planilha de
mentira que vários testes usam. Imita só o pedaço do xlwings que o código usa:

- ``ws.range(linha, coluna)`` devolve SEMPRE a mesma célula (escrever nela fica);
- ``ws.range((l1, c1), (l2, c2)).value`` devolve no formato do xlwings: escalar
  pra uma célula, lista simples pra uma linha OU uma coluna, lista de listas pro
  resto;
- ``celula.formula``: a fórmula, se houver; senão o próprio valor em texto (é o
  que o xlwings devolve pra uma constante digitada); ``""`` pra célula vazia;
- ``ws.used_range.last_cell`` é calculado na hora, a partir das células com
  valor — então um rodapé escrito pelo código "aumenta" a aba, como no Excel;
- ``wb.save()``/``wb.close()`` só contam quantas vezes foram chamados.

``ExcelFalso`` faz o papel de ``excel_app.app_excel`` + ``abrir_livro``: todo
caminho aberto devolve o MESMO livro em memória e a abertura fica registrada
(caminho, somente leitura?). Arquivo em disco, quando existe no teste, é só
marcador de "existe" pra transação/``Path.exists`` — o conteúdo de verdade é o
livro em memória.
"""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

from treinamentos_its.config import ABAS_DE_CARGO

NOME_MATRIZ = "Matriz - Atualização"
NOME_VERSIONAMENTO = "Versionamento Mês"


class CelulaFalsa:
    def __init__(self, value=None, formula=None):
        self.value = value
        self._formula = formula
        self.color = None
        self.font = SimpleNamespace(color=None, bold=False, size=None)

    @property
    def formula(self):
        if self._formula is not None:
            return self._formula
        if self.value is None:
            return ""
        return str(self.value)

    @formula.setter
    def formula(self, nova):
        self._formula = nova


class IntervaloFalso:
    def __init__(self, aba, inicio, fim):
        self.aba = aba
        self.inicio = inicio
        self.fim = fim

    @property
    def value(self):
        (l1, c1), (l2, c2) = self.inicio, self.fim
        grade = [
            [self.aba.celula(linha, coluna).value for coluna in range(c1, c2 + 1)]
            for linha in range(l1, l2 + 1)
        ]
        if l1 == l2 and c1 == c2:
            return grade[0][0]
        if l1 == l2:
            return grade[0]
        if c1 == c2:
            return [linha[0] for linha in grade]
        return grade


class AbaFalsa:
    def __init__(self, name, valores=None, formulas=None):
        self.name = name
        self.cells: dict[tuple[int, int], CelulaFalsa] = {}
        self.api = SimpleNamespace(AutoFilterMode=False)
        for (linha, coluna), valor in (valores or {}).items():
            self.celula(linha, coluna).value = valor
        for (linha, coluna), formula in (formulas or {}).items():
            self.celula(linha, coluna).formula = formula

    def celula(self, linha, coluna) -> CelulaFalsa:
        return self.cells.setdefault((linha, coluna), CelulaFalsa())

    def range(self, inicio, fim=None):
        if isinstance(inicio, tuple):
            return IntervaloFalso(self, inicio, fim if fim is not None else inicio)
        return self.celula(inicio, fim)

    @property
    def used_range(self):
        ocupadas = [
            chave for chave, celula in self.cells.items() if celula.value not in (None, "")
        ]
        linha = max((l for l, _c in ocupadas), default=1)
        coluna = max((c for _l, c in ocupadas), default=1)
        return SimpleNamespace(last_cell=SimpleNamespace(row=linha, column=coluna))

    def coluna_a(self) -> list:
        """Os textos da coluna A, de cima pra baixo (o rodapé mora aqui)."""
        ultima = self.used_range.last_cell.row
        return [self.celula(linha, 1).value for linha in range(1, ultima + 1)]

    def estado(self) -> dict:
        """Fotografia de valor + cor de todas as células com conteúdo."""
        return {
            chave: (celula.value, celula.color, celula.font.color)
            for chave, celula in self.cells.items()
            if celula.value not in (None, "")
        }


class ColecaoAbas(list):
    def __getitem__(self, chave):
        if isinstance(chave, str):
            for aba in self:
                if aba.name == chave:
                    return aba
            raise KeyError(chave)
        return super().__getitem__(chave)


class LivroFalso:
    def __init__(self, *abas, name="Planos e Macs Setembro 2026.xlsx", autosave=True):
        self.sheets = ColecaoAbas(abas)
        self.name = name
        self.api = SimpleNamespace(AutoSaveOn=autosave, ReadOnly=False)
        self.salvamentos = 0
        self.fechamentos = 0

    def save(self):
        self.salvamentos += 1

    def close(self):
        self.fechamentos += 1


class ExcelFalso:
    """``app_excel`` + ``abrir_livro`` de mentira; todo arquivo é o mesmo livro."""

    def __init__(self, livro: LivroFalso):
        self.livro = livro
        self.aberturas: list[tuple[Path, bool]] = []

    @contextmanager
    def app_excel(self, visible=False):
        yield self

    def abrir_livro(self, _app, caminho, *, read_only=False):
        self.aberturas.append((Path(caminho), read_only))
        return self.livro

    def instalar(self, monkeypatch, *modulos) -> "ExcelFalso":
        for modulo in modulos:
            monkeypatch.setattr(modulo, "app_excel", self.app_excel)
            monkeypatch.setattr(modulo, "abrir_livro", self.abrir_livro)
        return self


# --- montadores ---------------------------------------------------------------


def coluna_do_codigo(indice: int) -> int:
    """Coluna "Planejado" do ``indice``-ésimo curso do bloco (C, E, G, ...)."""
    return 3 + 2 * indice


def aba_de_cargo(
    nome: str,
    pessoas: list[tuple[str, dict[str, tuple[object, object]]]],
    codigos: tuple[str, ...] = ("DOC-POP-0000001",),
    *,
    linha_titulo: int = 1,
    rodape: tuple[str, ...] = (),
    aba: AbaFalsa | None = None,
) -> AbaFalsa:
    """Aba de cargo com UM bloco 'Treinamentos' no layout que ``_ler_blocos`` lê.

    - ``linha_titulo``: 'Treinamentos' na coluna A;
    - ``+1``: os códigos (C, E, G, ...), coluna do "Planejado"; o "Executado"
      (status) é a coluna da direita;
    - ``+2``: a linha "Versão" (fórmula sobre a Matriz; o código nunca escreve
      aqui);
    - ``+4`` em diante: uma pessoa por linha, nome na coluna B.

    ``pessoas``: ``[(nome, {codigo: (planejado, status)})]``. ``rodape``: linhas
    de texto já escritas na coluna A depois do bloco (herança do mês anterior).
    Passe ``aba`` pra acrescentar mais um bloco numa aba já montada.
    """
    aba = aba or AbaFalsa(nome)
    aba.celula(linha_titulo, 1).value = "Treinamentos"
    for indice, codigo in enumerate(codigos):
        coluna = coluna_do_codigo(indice)
        aba.celula(linha_titulo + 1, coluna).value = codigo
        aba.celula(linha_titulo + 2, coluna).value = 1.0
        aba.celula(linha_titulo + 2, coluna).formula = (
            f"=VLOOKUP(\"{codigo}\",'Matriz - Atualização'!$C:$E,3,0)"
        )
        aba.celula(linha_titulo + 3, coluna).value = "Planejado"
        aba.celula(linha_titulo + 3, coluna + 1).value = "Executado"
    aba.celula(linha_titulo + 2, 2).value = "Versão"
    linha = linha_titulo + 4
    for nome_pessoa, cursos in pessoas:
        aba.celula(linha, 2).value = nome_pessoa
        for indice, codigo in enumerate(codigos):
            planejado, status = cursos.get(codigo, (None, "OK"))
            coluna = coluna_do_codigo(indice)
            aba.celula(linha, coluna).value = planejado
            aba.celula(linha, coluna + 1).value = status
        linha += 1
    linha += 1
    for texto in rodape:
        aba.celula(linha, 1).value = texto
        linha += 1
    return aba


def linha_da_pessoa(aba: AbaFalsa, nome: str) -> int:
    for (linha, coluna), celula in aba.cells.items():
        if coluna == 2 and celula.value == nome:
            return linha
    raise KeyError(nome)


def celulas_da_pessoa(aba: AbaFalsa, nome: str, indice_codigo: int = 0):
    """``(célula Planejado, célula Executado)`` da pessoa no curso indicado."""
    linha = linha_da_pessoa(aba, nome)
    coluna = coluna_do_codigo(indice_codigo)
    return aba.celula(linha, coluna), aba.celula(linha, coluna + 1)


def aba_matriz(
    linhas: list[tuple[object, object, object, object]], *, inicio: int = 5
) -> AbaFalsa:
    """``linhas``: ``(código C, E, F, fórmula de F ou None)`` a partir da linha 5.

    Fórmula ``None`` com F preenchida = valor digitado por cima da fórmula (o
    ``formula`` do dublê devolve o valor, como o xlwings).
    """
    aba = AbaFalsa(NOME_MATRIZ)
    aba.celula(4, 3).value = "Código"
    aba.celula(4, 5).value = "Versão atual"
    aba.celula(4, 6).value = "Versão nova"
    for deslocamento, (codigo, atual, nova, formula) in enumerate(linhas):
        linha = inicio + deslocamento
        aba.celula(linha, 3).value = codigo
        aba.celula(linha, 5).value = atual
        aba.celula(linha, 6).value = nova
        if formula is not None:
            aba.celula(linha, 6).formula = formula
    return aba


FORMULA_F = "=SEERRO(PROCV(C5;'Versionamento Mês'!$A:$D;4;0);\"Verificar\")"


def aba_versionamento(datas: dict[str, object]) -> AbaFalsa:
    """Aba 'Versionamento Mês' com ``{código: data de aprovação}``."""
    aba = AbaFalsa(NOME_VERSIONAMENTO)
    for coluna, titulo in enumerate(
        ("Document Number", "Title", "Approved Date", "Version"), start=1
    ):
        aba.celula(1, coluna).value = titulo
    for linha, (codigo, data) in enumerate(datas.items(), start=2):
        aba.celula(linha, 1).value = codigo
        aba.celula(linha, 2).value = f"Título de {codigo}"
        aba.celula(linha, 3).value = data
        aba.celula(linha, 4).value = 2.0
    return aba


def abas_de_cargo_vazias(exceto: tuple[str, ...] = ()) -> list[AbaFalsa]:
    """As abas de cargo reais do ``config`` que o teste não preencheu (vazias)."""
    return [AbaFalsa(nome) for nome in ABAS_DE_CARGO if nome not in exceto]
