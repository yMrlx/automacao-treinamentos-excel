"""Achados do QA de "operadora errando" e da revisão do Codex (05/10/2026).

- versão colada como NÚMERO (Excel converte '1.0' em 1) -> Matriz toda vermelha;
- planilha aberta no Excel de alguém -> falso RESTAURAÇÃO INCOMPLETA (código 3);
- caixa de pergunta prometendo "não grava" quando só o arquivo existia;
- falha ao matar o Excel passando calada.
Nada aqui abre Excel nem resolve o PATH_BASE real.
"""

import logging
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

from treinamentos_its import backup, cli, conclusao, excel_app
from treinamentos_its.plano_execucao import (
    EXPLICACAO_REPROCESSAR,
    EXPLICACAO_REPROCESSAR_SEM_RODAPE,
    texto_da_pergunta,
)
from treinamentos_its.vigentes import bloco_para_colar, montar_bloco_canonico, versao_como_texto

# --------------------------------------------------------------- versão como texto


@pytest.mark.parametrize(
    ("valor", "esperado"),
    [(1, "1.0"), (12.0, "12.0"), ("3.0", "3.0"), ("3.0.0", "3.0.0"), (1.1, 1.1),
     (None, None), (True, True), ("VERIFICAR", "VERIFICAR")],
)
def test_versao_como_texto(valor, esperado):
    assert versao_como_texto(valor) == esperado


def test_bloco_canonico_grava_versao_numerica_como_texto_n_ponto_0():
    cabecalho = ["Document Number", "Title", "Status", "Version", "Owning Department",
                 "Impacted Departments", "Approved Date", "Approver"]
    linhas = [["DOC-POP-1", "t", "Em vigor", 4, "d", "d", datetime(2026, 8, 1), "x"],
              ["DOC-POP-2", "t", "Em vigor", "2.0", "d", "d", datetime(2026, 8, 2), "x"]]
    bloco, _avisos = montar_bloco_canonico(cabecalho, linhas)
    assert [linha[3] for linha in bloco[1:]] == ["4.0", "2.0"]


def test_bloco_para_colar_protege_texto_e_nao_mexe_no_resto():
    data = datetime(2026, 8, 11)
    bloco = [["Document Number", "Version"], ["DOC-POP-1", "1.0", data, 7, None, ""]]
    colar = bloco_para_colar(bloco)
    assert colar == [["'Document Number", "'Version"], ["'DOC-POP-1", "'1.0", data, 7, None, ""]]
    assert bloco[1][1] == "1.0"  # o original (usado na conferência) não muda


# ------------------------------------------------- planilha aberta no Excel de alguém


def test_arquivos_em_uso_detecta_o_travado(monkeypatch, tmp_path):
    livre = tmp_path / "livre.xlsx"
    travado = tmp_path / "travado.xlsx"
    livre.write_bytes(b"x")
    travado.write_bytes(b"x")
    abrir_de_verdade = open

    def abrir(caminho, modo="r", *args, **kwargs):
        if Path(caminho) == travado and "+" in modo:
            raise PermissionError(13, "Permission denied")
        return abrir_de_verdade(caminho, modo, *args, **kwargs)

    monkeypatch.setattr(backup, "open", abrir, raising=False)
    assert backup.arquivos_em_uso([livre, travado, tmp_path / "nao_existe.xlsx"]) == [travado]
    assert livre.read_bytes() == b"x" and travado.read_bytes() == b"x"


def test_ciclo_com_planilha_aberta_para_antes_de_mexer(monkeypatch, tmp_path, caplog):
    from treinamentos_its import config

    monkeypatch.setattr(config, "PASTA_PLANOS_MACS", tmp_path)
    monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _m: [])
    monkeypatch.setattr(
        cli, "arquivos_em_uso", lambda caminhos: [Path("Planos e Macs Agosto 2026.xlsx")]
    )

    def _nao_deveria(_meses):
        raise AssertionError("abriu a transação com a planilha aberta")

    monkeypatch.setattr(cli, "_transacao_do_ciclo", _nao_deveria)
    with caplog.at_level(logging.INFO, logger="treinamentos_its"):
        assert cli._run_pipeline_mensal(8, 2026) == 1
    assert "está ABERTO no Excel" in caplog.text
    assert "Nada foi alterado" in caplog.text


def test_concluir_com_planilha_aberta_sai_1_sem_restauracao(monkeypatch, tmp_path, caplog):
    arquivo = tmp_path / "Planos e Macs Setembro 2026.xlsx"
    arquivo.write_bytes(b"setembro")
    monkeypatch.setattr(conclusao, "PASTA_PLANOS_MACS", tmp_path)
    ocorrencia = conclusao.OcorrenciaPlanos("Estagiario", 124, 5, 120, "ATRASADO")
    monkeypatch.setattr(conclusao, "_localizar_ocorrencias", lambda *_a: [ocorrencia])
    monkeypatch.setattr(conclusao, "arquivos_em_uso", lambda caminhos: list(caminhos))

    def _nao_deveria(*_a, **_k):
        raise AssertionError("abriu a transação com a planilha aberta")

    monkeypatch.setattr(conclusao, "TransacaoArquivo", _nao_deveria)
    with caplog.at_level(logging.INFO, logger="treinamentos_its"):
        codigo = conclusao.registrar_conclusao(
            "Carlos Andrade", "DOC-POP-0000576", datetime(2026, 9, 30)
        )
    assert codigo == 1
    assert "está ABERTO no Excel" in caplog.text
    assert "RESTAURAÇÃO" not in caplog.text
    assert arquivo.read_bytes() == b"setembro"


# ------------------------------------------------- caixa de pergunta fiel ao backend


def _log(nivel, mensagem):
    return f"05/10/2026 14:26:06  {nivel:<8}  {mensagem}"


def test_pergunta_sem_rodape_nao_promete_que_nada_e_gravado():
    linhas = [
        _log("WARNING", "O ciclo de Setembro/2026 parece JÁ ter sido processado:"),
        _log("WARNING", "  - a planilha 'Planos e Macs Setembro 2026.xlsx' de destino já existe"),
        "Rodar mesmo assim? (s/n): ",
    ]
    texto = texto_da_pergunta(linhas)
    assert texto.endswith(EXPLICACAO_REPROCESSAR_SEM_RODAPE)
    assert EXPLICACAO_REPROCESSAR not in texto


def test_pergunta_com_rodape_mantem_a_explicacao_do_nada_novo():
    linhas = [
        _log("WARNING", "O ciclo de Setembro/2026 parece JÁ ter sido processado:"),
        _log("WARNING", "  - o rodapé 'Referente Agosto 2026' já está escrito em: Analista"),
        "Rodar mesmo assim? (s/n): ",
    ]
    assert texto_da_pergunta(linhas).endswith(EXPLICACAO_REPROCESSAR)


# ------------------------------------------------- Excel que não morre não passa calado


def test_kill_que_falha_depois_de_quit_que_falhou_vira_erro_com_pid(caplog):
    def _quit():
        raise RuntimeError("Excel ocupado")

    def _kill():
        raise PermissionError("acesso negado")

    app = SimpleNamespace(quit=_quit, kill=_kill, impl=SimpleNamespace(_pid=4321))
    with caplog.at_level(logging.WARNING, logger="treinamentos_its"):
        excel_app._encerrar_excel(app)
    assert "NÃO consegui encerrar o Excel invisível (processo 4321)" in caplog.text


def test_kill_que_falha_depois_de_quit_bom_continua_silencioso(caplog):
    def _kill():
        raise OSError("processo já saiu")

    app = SimpleNamespace(quit=lambda: None, kill=_kill, impl=SimpleNamespace(_pid=4321))
    with caplog.at_level(logging.WARNING, logger="treinamentos_its"):
        excel_app._encerrar_excel(app)
    assert caplog.text == ""


@pytest.mark.parametrize(
    ("valor", "esperado"),
    [("4.0", "'4.0"), (4, "'4.0"), (4.0, "'4.0"), ("3.0.0", "'3.0.0"), (1.1, 1.1), (None, None)],
)
def test_versao_para_gravar_vai_como_texto_protegido(valor, esperado):
    from treinamentos_its.excel_utils import versao_para_gravar

    assert versao_para_gravar(valor) == esperado


# ------------------------------------------- revisão do Codex de 06/10/2026


def test_fechar_o_excel_limita_as_repeticoes_do_xlwings_e_devolve_o_padrao(monkeypatch):
    import sys
    import types

    xlwindows = types.ModuleType("xlwings._xlwindows")
    xlwindows.N_COM_ATTEMPTS = 0  # padrão do xlwings: infinito
    pacote = types.ModuleType("xlwings")
    pacote._xlwindows = xlwindows
    monkeypatch.setitem(sys.modules, "xlwings", pacote)
    monkeypatch.setitem(sys.modules, "xlwings._xlwindows", xlwindows)
    vistos = []
    app = SimpleNamespace(
        quit=lambda: vistos.append(xlwindows.N_COM_ATTEMPTS), kill=lambda: None,
        impl=SimpleNamespace(_pid=1),
    )

    excel_app._encerrar_excel(app)

    assert vistos == [excel_app.TENTATIVAS_COM_AO_FECHAR] and vistos[0] > 0
    assert xlwindows.N_COM_ATTEMPTS == 0


def test_painel_nao_diz_sem_erros_quando_o_excel_ficou_vivo():
    from treinamentos_its import plano_execucao as pe

    log = [
        "06/10/2026 10:00:00  INFO      Planos e Macs atualizado com sucesso.",
        "06/10/2026 10:00:01  ERROR     NÃO consegui encerrar o Excel invisível (processo 7): x",
    ]
    assert pe.desfecho_da_execucao(pe.ACAO_SYNC, 0, log) == (
        pe.TEXTO_EXCEL_NAO_ENCERRADO, pe.NIVEL_ALERTA
    )
    assert pe.desfecho_da_execucao(pe.ACAO_SYNC, 1, log)[1] == pe.NIVEL_ERRO
