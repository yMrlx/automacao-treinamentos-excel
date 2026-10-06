"""Testes do ``excel_app.abrir_livro`` com dublê do xlwings — sem Excel.

O furo que isto fecha: com os alertas desligados, um arquivo aberto em outro
computador/Excel abre SOMENTE LEITURA em silêncio, e quem ia escrever não
percebia nada. Agora quem abre pra escrita recebe um erro claro na hora.
"""

from types import SimpleNamespace

import pytest

from treinamentos_its import excel_app


class _LivroFalso:
    def __init__(self, somente_leitura=False, api_quebrada=False):
        self.fechado = False
        if api_quebrada:
            self._api = None
        else:
            self._api = SimpleNamespace(ReadOnly=somente_leitura)

    @property
    def api(self):
        if self._api is None:
            raise RuntimeError("COM desconectado")
        return self._api

    def close(self):
        self.fechado = True


class _AppFalso:
    def __init__(self, livro):
        self.livro = livro
        self.aberturas = []
        self.books = SimpleNamespace(open=self._abrir)

    def _abrir(self, caminho, **kwargs):
        self.aberturas.append((caminho, kwargs))
        return self.livro


def test_abrir_pra_escrita_normal_devolve_o_livro():
    livro = _LivroFalso(somente_leitura=False)
    app = _AppFalso(livro)

    assert excel_app.abrir_livro(app, r"C:\x\Procedimentos - Time de TI.xlsx") is livro
    assert livro.fechado is False
    assert app.aberturas == [
        (r"C:\x\Procedimentos - Time de TI.xlsx", {"update_links": False, "read_only": False})
    ]


def test_abrir_pra_escrita_e_vir_somente_leitura_fecha_e_recusa():
    livro = _LivroFalso(somente_leitura=True)
    app = _AppFalso(livro)

    with pytest.raises(excel_app.ArquivoAbertoSomenteLeitura) as erro:
        excel_app.abrir_livro(app, r"C:\x\Procedimentos - Time de TI.xlsx")

    mensagem = str(erro.value)
    assert "'Procedimentos - Time de TI.xlsx' abriu SOMENTE LEITURA" in mensagem
    assert "aberto em outro computador/Excel" in mensagem
    assert "Feche e rode de novo" in mensagem
    assert livro.fechado is True


def test_nao_conseguir_ler_o_readonly_tambem_recusa():
    livro = _LivroFalso(api_quebrada=True)

    with pytest.raises(excel_app.ArquivoAbertoSomenteLeitura, match="Não consegui confirmar"):
        excel_app.abrir_livro(_AppFalso(livro), "Planos e Macs Setembro 2026.xlsx")
    assert livro.fechado is True


def test_fechar_que_falha_nao_esconde_o_motivo_real():
    livro = _LivroFalso(somente_leitura=True)

    def _fechar_quebrado():
        raise RuntimeError("Excel travou")

    livro.close = _fechar_quebrado

    with pytest.raises(excel_app.ArquivoAbertoSomenteLeitura, match="SOMENTE LEITURA"):
        excel_app.abrir_livro(_AppFalso(livro), "Planos e Macs Setembro 2026.xlsx")


def test_abrir_somente_leitura_de_proposito_nao_confere_nada():
    # Fonte que só é lida (read_only=True): abrir read-only é o esperado.
    livro = _LivroFalso(somente_leitura=True)
    app = _AppFalso(livro)

    assert excel_app.abrir_livro(app, "backup.xlsx", read_only=True) is livro
    assert livro.fechado is False
    assert app.aberturas[0][1] == {"update_links": False, "read_only": True}


def test_erro_e_runtimeerror_pra_quem_ja_captura_exception():
    # sync/planos_macs/conclusao capturam Exception e fazem a transação desfazer.
    assert issubclass(excel_app.ArquivoAbertoSomenteLeitura, RuntimeError)


# --- app_excel sempre encerra o processo (05/10/2026: Excel preso em diálogo) ---


class _ProcessoExcelFalso:
    def __init__(self, visible=False, *, quit_explode=None):
        self.chamadas: list[str] = []
        self._quit_explode = quit_explode
        self.display_alerts = True
        self.screen_updating = True

    def quit(self):
        self.chamadas.append("quit")
        if self._quit_explode is not None:
            raise self._quit_explode

    def kill(self):
        self.chamadas.append("kill")


def _instalar_xlwings(monkeypatch, app):
    import sys

    monkeypatch.setitem(sys.modules, "xlwings", SimpleNamespace(App=lambda visible=False: app))


def test_app_excel_fecha_e_encerra_o_processo(monkeypatch):
    app = _ProcessoExcelFalso()
    _instalar_xlwings(monkeypatch, app)

    with excel_app.app_excel() as aberto:
        assert aberto is app
        assert app.display_alerts is False and app.screen_updating is False

    assert app.chamadas == ["quit", "kill"]


def test_excel_travado_no_quit_ainda_e_morto_e_avisa(monkeypatch, caplog):
    import logging

    app = _ProcessoExcelFalso(quit_explode=RuntimeError("OLE error 0x800ac472"))
    _instalar_xlwings(monkeypatch, app)

    with caplog.at_level(logging.WARNING, logger="treinamentos_its"):
        with excel_app.app_excel():
            pass

    assert app.chamadas == ["quit", "kill"]
    assert "não fechou normalmente" in caplog.text


def test_erro_ao_fechar_nao_esconde_o_erro_da_operacao(monkeypatch):
    app = _ProcessoExcelFalso(quit_explode=RuntimeError("Excel ocupado"))
    _instalar_xlwings(monkeypatch, app)

    with pytest.raises(ValueError, match="erro real"):
        with excel_app.app_excel():
            raise ValueError("erro real da importação")

    assert app.chamadas == ["quit", "kill"]
