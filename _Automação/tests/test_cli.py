from treinamentos_its import cli, config
from treinamentos_its.dominio import AvaliacaoMatriz


MESES = {
    "anterior": "Agosto",
    "ano_ant_relativo": 2026,
    "atual": "Setembro",
    "ano_atual": 2026,
}


def _preparar(monkeypatch, tmp_path, *, atualiza=True):
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    (pasta / "Planos e Macs Agosto 2026.xlsx").write_text("origem", encoding="utf-8")
    monkeypatch.setattr(config, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(config, "PASTA_PRE_COLAGEM", pasta / "_pre-colagem")
    monkeypatch.setattr(cli, "obter_meses", lambda *_: MESES)
    monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _meses: [])
    monkeypatch.setattr(
        cli,
        "avaliar_arquivo_matriz",
        lambda _arquivo: AvaliacaoMatriz(),
    )
    monkeypatch.setattr(cli, "atualizar_planos_e_macs", lambda *_args: atualiza)
    return pasta


def test_pipeline_altera_somente_o_planos_novo(monkeypatch, tmp_path):
    pasta = _preparar(monkeypatch, tmp_path)
    assert cli._run_pipeline_mensal(9, 2026) == 0
    assert (pasta / "Planos e Macs Setembro 2026.xlsx").exists()
    assert not (tmp_path / "Procedimentos").exists()


def test_falha_remove_planos_que_nasceu_no_ciclo(monkeypatch, tmp_path):
    pasta = _preparar(monkeypatch, tmp_path, atualiza=False)
    assert cli._run_pipeline_mensal(9, 2026) == 1
    assert not (pasta / "Planos e Macs Setembro 2026.xlsx").exists()


def test_guarda_recusada_nao_cria_destino(monkeypatch, tmp_path):
    pasta = _preparar(monkeypatch, tmp_path)
    monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _meses: ["já existe"])
    monkeypatch.setattr(cli, "perguntar_sim_nao", lambda _pergunta: False)
    assert cli._run_pipeline_mensal(9, 2026) == 0
    assert not (pasta / "Planos e Macs Setembro 2026.xlsx").exists()


def test_parser_recusa_abreviacao_da_flag_teste():
    parser = cli.construir_parser()
    try:
        parser.parse_args(["sync", "--tes"])
    except SystemExit as erro:
        assert erro.code == 2
    else:
        raise AssertionError("--tes não poderia ser aceito")


def test_parser_aceita_contratos_da_gui():
    parser = cli.construir_parser()
    assert parser.parse_args(["sync", "--mes", "8", "--ano", "2026", "--teste"]).comando == "sync"
    assert parser.parse_args(["conferir", "--pops-vigentes", "x.xlsx"]).comando == "conferir"
    assert parser.parse_args(["concluir", "Carla", "DOC-POP-1", "--data", "25/09/2026"]).comando == "concluir"


# =============================================================================
# Cobertura recriada depois do a5dad36: a transação do ciclo (só o Planos e Macs
# do mês novo) e o `main()`. Arquivos de verdade numa pasta temporária; as
# etapas que abririam o Excel são dublês que escrevem no arquivo como o Excel
# escreveria (AutoSave incluso).
# =============================================================================

import logging
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

from treinamentos_its import ambiente_teste, backup


@pytest.fixture
def ciclo(monkeypatch, tmp_path):
    """Pasta do Planos e Macs em ``tmp_path``. Nunca resolve o PATH_BASE real."""
    pasta = tmp_path / "Planos e Macs"
    pasta.mkdir()
    origem = pasta / "Planos e Macs Agosto 2026.xlsx"
    origem.write_text("agosto - entregável fechado", encoding="utf-8")
    monkeypatch.setattr(config, "PASTA_PLANOS_MACS", pasta)
    monkeypatch.setattr(config, "PASTA_PRE_COLAGEM", pasta / "_pre-colagem")
    monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _meses: [])
    monkeypatch.setattr(cli, "avaliar_arquivo_matriz", lambda _arquivo: AvaliacaoMatriz())
    return SimpleNamespace(
        pasta=pasta,
        origem=origem,
        destino=pasta / "Planos e Macs Setembro 2026.xlsx",
        copias=pasta / "_pre-colagem",
    )


def _planos_que_escreve(*, resultado=True, explode=None):
    """Dublê do `atualizar_planos_e_macs`: escreve no arquivo e depois decide."""

    def _atualizar(_avaliacao, _meses, arquivo):
        Path(arquivo).write_text("setembro DEPOIS do ciclo", encoding="utf-8")
        if explode is not None:
            raise explode
        return resultado

    return _atualizar


def _copias(ciclo) -> dict[str, str]:
    return {
        copia.name: copia.read_text(encoding="utf-8")
        for copia in ciclo.copias.glob("*.xlsx")
    }


def _ler(arquivo: Path) -> str:
    return arquivo.read_text(encoding="utf-8")


class TestTransacaoDoCiclo:
    def test_restauracao_que_falha_sai_com_3_e_diz_de_onde_restaurar(
        self, monkeypatch, ciclo, caplog
    ):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _planos_que_escreve(resultado=False))
        monkeypatch.setattr(backup, "restaurar_arquivo", lambda *_a, **_k: False)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            codigo = cli._run_pipeline_mensal(9, 2026)

        assert codigo == backup.CODIGO_SAIDA_RESTAURACAO_INCOMPLETA == 3
        [copia] = ciclo.copias.glob("*.xlsx")
        assert _ler(copia) == "setembro ANTES"
        assert "RESTAURAÇÃO INCOMPLETA" in caplog.text
        assert (
            f"'{ciclo.destino}': com o Excel FECHADO, copie '{copia}' por cima dele"
            in caplog.text
        )
        assert "do mês novo voltou ao estado anterior" not in caplog.text
        assert _ler(ciclo.destino) == "setembro DEPOIS do ciclo"
        assert _ler(ciclo.origem) == "agosto - entregável fechado"

    def test_excecao_com_restauracao_que_falha_tambem_sai_com_3(
        self, monkeypatch, ciclo, caplog
    ):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(
            cli,
            "atualizar_planos_e_macs",
            _planos_que_escreve(explode=RuntimeError("Excel caiu")),
        )
        monkeypatch.setattr(backup, "restaurar_arquivo", lambda *_a, **_k: False)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 3
        assert "Erro inesperado no ciclo" in caplog.text
        assert "RESTAURAÇÃO INCOMPLETA" in caplog.text

    def test_arquivo_novo_que_nao_pode_ser_apagado_sai_com_3(
        self, monkeypatch, ciclo, caplog
    ):
        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _planos_que_escreve(resultado=False))
        unlink_original = Path.unlink

        def _unlink_travado(self, *args, **kwargs):
            if self == ciclo.destino:
                raise PermissionError("aberto no Excel")
            return unlink_original(self, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", _unlink_travado)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 3
        assert (
            f"'{ciclo.destino}': nasceu nesta execução e ficou pela metade - apague-o à mão"
            in caplog.text
        )
        assert ciclo.destino.exists()

    @pytest.mark.parametrize("ja_existia", [False, True], ids=["arquivo novo", "reprocessamento"])
    @pytest.mark.parametrize("onde", ["na validação", "no Planos e Macs"])
    def test_excecao_no_meio_do_ciclo_desfaz(self, monkeypatch, ciclo, caplog, ja_existia, onde):
        if ja_existia:
            ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        if onde == "na validação":
            def _avaliar_explode(arquivo):
                Path(arquivo).write_text("lixo do AutoSave", encoding="utf-8")
                raise RuntimeError("Excel caiu")

            monkeypatch.setattr(cli, "avaliar_arquivo_matriz", _avaliar_explode)
        monkeypatch.setattr(
            cli,
            "atualizar_planos_e_macs",
            _planos_que_escreve(explode=RuntimeError("Excel caiu")),
        )

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 1

        if ja_existia:
            assert _ler(ciclo.destino) == "setembro ANTES"
        else:
            assert not ciclo.destino.exists()
        assert _ler(ciclo.origem) == "agosto - entregável fechado"
        assert "Erro inesperado no ciclo" in caplog.text
        assert "do mês novo voltou ao estado anterior" in caplog.text
        assert "RESTAURAÇÃO INCOMPLETA" not in caplog.text

    def test_reprocessamento_com_falha_restaura_o_arquivo_do_mes(self, monkeypatch, ciclo):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _m: ["já existe"])
        monkeypatch.setattr(cli, "perguntar_sim_nao", lambda _pergunta: True)
        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _planos_que_escreve(resultado=False))

        assert cli._run_pipeline_mensal(9, 2026) == 1

        assert _ler(ciclo.destino) == "setembro ANTES"
        assert list(_copias(ciclo).values()) == ["setembro ANTES"]
        assert _ler(ciclo.origem) == "agosto - entregável fechado"

    def test_reprocessamento_bem_sucedido_mantem_o_novo_e_deixa_a_copia(
        self, monkeypatch, ciclo
    ):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _m: ["já existe"])
        monkeypatch.setattr(cli, "perguntar_sim_nao", lambda _pergunta: True)
        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _planos_que_escreve())

        assert cli._run_pipeline_mensal(9, 2026) == 0

        assert _ler(ciclo.destino) == "setembro DEPOIS do ciclo"
        assert list(_copias(ciclo).values()) == ["setembro ANTES"]

    def test_falha_com_restauracao_ok_avisa_depois_de_restaurar(
        self, monkeypatch, ciclo, caplog
    ):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _planos_que_escreve(resultado=False))

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 1

        mensagens = [registro.getMessage() for registro in caplog.records]
        restaurado = next(i for i, m in enumerate(mensagens) if "RESTAURADO" in m)
        desfeito = next(i for i, m in enumerate(mensagens) if "do mês novo voltou ao estado anterior" in m)
        assert restaurado < desfeito
        assert "RESTAURAÇÃO INCOMPLETA" not in caplog.text

    def test_sem_copia_de_restauracao_nada_roda(self, monkeypatch, ciclo, caplog):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        chamadas: list[str] = []
        monkeypatch.setattr(
            cli, "avaliar_arquivo_matriz", lambda _a: chamadas.append("avaliar")
        )

        def _copia_quebrada(*_args, **_kwargs):
            raise PermissionError("disco cheio")

        monkeypatch.setattr(backup.shutil, "copy2", _copia_quebrada)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 1

        assert chamadas == []
        assert _ler(ciclo.destino) == "setembro ANTES"
        assert "Sem cópia de restauração" in caplog.text
        assert "RESTAURAÇÃO INCOMPLETA" not in caplog.text

    @pytest.mark.parametrize(
        "avaliacao",
        [None, AvaliacaoMatriz(bloqueios=["regressão de versão: QU (2.0 -> 1.0)"])],
        ids=["leitura falhou", "trava"],
    )
    def test_validacao_que_bloqueia_nao_chama_o_planos_e_desfaz(
        self, monkeypatch, ciclo, caplog, avaliacao
    ):
        chamadas: list[str] = []
        monkeypatch.setattr(cli, "avaliar_arquivo_matriz", lambda _a: avaliacao)
        monkeypatch.setattr(
            cli, "atualizar_planos_e_macs", lambda *_a: chamadas.append("planos") or True
        )

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 1

        assert chamadas == []
        assert not ciclo.destino.exists()
        assert "bloqueou o ciclo" in caplog.text

    def test_ctrl_c_no_meio_do_ciclo_tambem_desfaz(self, monkeypatch, ciclo):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(
            cli, "atualizar_planos_e_macs", _planos_que_escreve(explode=KeyboardInterrupt())
        )

        with pytest.raises(KeyboardInterrupt):
            cli._run_pipeline_mensal(9, 2026)

        assert _ler(ciclo.destino) == "setembro ANTES"


# --- main(): o que roda e com que código sai -------------------------------------


def _proibir_tudo(monkeypatch) -> None:
    def _nao_deveria(*_args, **_kwargs):
        raise AssertionError("nada deveria ter rodado")

    for nome in ("_run_pipeline_mensal", "conferir_matriz", "registrar_conclusao"):
        monkeypatch.setattr(cli, nome, _nao_deveria)


def _main(monkeypatch, *argumentos) -> int:
    monkeypatch.setattr("sys.argv", ["main.py", *argumentos])
    with pytest.raises(SystemExit) as saida:
        cli.main()
    return saida.value.code


ARGV_COM_TESTE = {
    "sync": ["sync", "--teste"],
    "conferir": ["conferir", "--teste"],
    "concluir": ["concluir", "Carla", "DOC-POP-1", "--teste"],
}


class TestMain:
    def test_sem_subcomando_nao_roda_o_ciclo(self, monkeypatch, caplog):
        # R10: `python main.py` sozinho caía no pipeline e mexia em PRODUÇÃO.
        _proibir_tudo(monkeypatch)
        assert _main(monkeypatch) == 2
        assert "Nenhum comando informado" in caplog.text

    def test_so_com_flags_de_referencia_tambem_nao_roda(self, monkeypatch):
        _proibir_tudo(monkeypatch)
        assert _main(monkeypatch, "--mes", "3", "--ano", "2026") == 2

    @pytest.mark.parametrize("comando", list(ARGV_COM_TESTE))
    def test_teste_com_a_base_de_producao_aborta(self, monkeypatch, tmp_path, caplog, comando):
        _proibir_tudo(monkeypatch)
        monkeypatch.setattr(config, "PATH_BASE", tmp_path / "producao")
        monkeypatch.setattr(ambiente_teste, "PASTA_SANDBOX", tmp_path / "_ambiente-teste")

        assert _main(monkeypatch, *ARGV_COM_TESTE[comando]) == 1
        assert "Nada foi executado" in caplog.text

    def test_teste_com_a_base_no_sandbox_roda(self, monkeypatch, tmp_path):
        sandbox = tmp_path / "_ambiente-teste"
        sandbox.mkdir()
        monkeypatch.setattr(ambiente_teste, "PASTA_SANDBOX", sandbox)
        monkeypatch.setattr(config, "PATH_BASE", sandbox)
        chamadas: list[tuple] = []
        monkeypatch.setattr(
            cli, "_run_pipeline_mensal", lambda *args: chamadas.append(args) or 0
        )
        assert _main(monkeypatch, "sync", "--teste") == 0
        assert chamadas == [(None, None, None)]

    @pytest.mark.parametrize("codigo", [0, 1, 3])
    def test_concluir_repassa_o_codigo_de_saida(self, monkeypatch, codigo):
        chamadas: list[tuple] = []
        monkeypatch.setattr(
            cli, "registrar_conclusao", lambda *args: chamadas.append(args) or codigo
        )
        assert _main(monkeypatch, "concluir", "Carla", "POP-1", "--data", "10/09/2026") == codigo
        assert chamadas == [("Carla", "POP-1", datetime(2026, 9, 10), None, None)]

    @pytest.mark.parametrize(
        "extra",
        [
            ["--data", "2026-09-10"],
            ["--mes", "9"],
            ["--ano", "2026"],
            ["--mes", "13", "--ano", "2026"],
        ],
        ids=["data ISO", "mes sem ano", "ano sem mes", "mes 13"],
    )
    def test_concluir_recusa_argumentos_ruins_antes_de_rodar(self, monkeypatch, extra):
        _proibir_tudo(monkeypatch)
        assert _main(monkeypatch, "concluir", "Carla", "POP-1", *extra) == 2

    @pytest.mark.parametrize("prefixo", ["--te", "--tes", "--test"])
    @pytest.mark.parametrize("comando", list(ARGV_COM_TESTE))
    def test_prefixo_da_flag_teste_e_recusado(self, monkeypatch, comando, prefixo):
        # `--tes` já foi aceito e rodou em PRODUÇÃO achando que era teste.
        _proibir_tudo(monkeypatch)
        argumentos = [prefixo if arg == "--teste" else arg for arg in ARGV_COM_TESTE[comando]]
        assert _main(monkeypatch, *argumentos) == 2

    @pytest.mark.parametrize("deu_certo, codigo", [(True, 0), (False, 1)])
    def test_conferir_sai_conforme_o_resultado(self, monkeypatch, deu_certo, codigo):
        chamadas: list[tuple] = []
        monkeypatch.setattr(
            cli, "conferir_matriz", lambda *args: chamadas.append(args) or deu_certo
        )
        assert _main(
            monkeypatch, "conferir", "--mes", "8", "--ano", "2026",
            "--pops-vigentes", "v.xlsx", "--pops-obsoletos", "o.xlsx",
        ) == codigo
        assert chamadas == [(8, 2026, "v.xlsx", "o.xlsx")]

    def test_sync_repassa_mes_ano_e_vigentes(self, monkeypatch):
        chamadas: list[tuple] = []
        monkeypatch.setattr(
            cli, "_run_pipeline_mensal", lambda *args: chamadas.append(args) or 3
        )
        assert _main(
            monkeypatch, "sync", "--mes", "3", "--ano", "2026", "--vigentes", "v.xlsx"
        ) == 3
        assert chamadas == [(3, 2026, "v.xlsx")]


class TestReprocessarSemNadaNovo:
    """Decisão de negócio (2026-10-02): mês já fechado + nenhum versionamento novo
    = não grava nada (antes empilhava um 2º bloco "Não houveram alterações")."""

    RODAPE_DO_CICLO = ["Referente Setembro 2026", "Não houveram alterações nos POPs"]

    def _reprocessar(self, monkeypatch, ciclo, avaliacao, *, reimporta=False):
        ciclo.destino.write_text("setembro ANTES", encoding="utf-8")
        monkeypatch.setattr(cli, "sinais_de_reprocessamento", lambda _m: ["já existe"])
        monkeypatch.setattr(cli, "perguntar_sim_nao", lambda _pergunta: True)
        monkeypatch.setattr(cli, "avaliar_arquivo_matriz", lambda _a: avaliacao)
        chamadas: list[str] = []

        def _atualizar(_avaliacao, _meses, arquivo):
            chamadas.append("atualizar")
            Path(arquivo).write_text("setembro DEPOIS do ciclo", encoding="utf-8")
            return True

        monkeypatch.setattr(cli, "atualizar_planos_e_macs", _atualizar)
        if reimporta:
            def _preparar(_meses, _vigentes=None):
                ciclo.destino.write_text("setembro com vigentes REIMPORTADOS", encoding="utf-8")
                return ciclo.destino

            monkeypatch.setattr(cli, "_preparar_planos_novo", _preparar)
        return chamadas

    def test_rodape_ja_escrito_e_nada_novo_nao_grava_nada(self, monkeypatch, ciclo, caplog):
        avaliacao = AvaliacaoMatriz(linhas_rodape=list(self.RODAPE_DO_CICLO))
        chamadas = self._reprocessar(monkeypatch, ciclo, avaliacao)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert cli._run_pipeline_mensal(9, 2026) == 0

        assert chamadas == []
        assert _ler(ciclo.destino) == "setembro ANTES"
        assert "Nada novo para registrar" in caplog.text
        assert _ler(ciclo.origem) == "agosto - entregável fechado"
        # O painel mostra desfecho neutro, não "Terminou sem erros".
        from treinamentos_its import plano_execucao as pe

        assert pe.desfecho_da_execucao(
            pe.ACAO_SYNC, 0, caplog.text.splitlines()
        ) == (pe.TEXTO_NADA_ALTERADO, pe.NIVEL_NEUTRO)

    def test_reimportacao_de_vigentes_tambem_e_desfeita(self, monkeypatch, ciclo):
        avaliacao = AvaliacaoMatriz(linhas_rodape=list(self.RODAPE_DO_CICLO))
        chamadas = self._reprocessar(monkeypatch, ciclo, avaliacao, reimporta=True)

        assert cli._run_pipeline_mensal(9, 2026) == 0

        assert chamadas == []
        assert _ler(ciclo.destino) == "setembro ANTES"

    def test_restauracao_que_falha_no_nada_novo_sai_com_3(self, monkeypatch, ciclo, caplog):
        avaliacao = AvaliacaoMatriz(linhas_rodape=list(self.RODAPE_DO_CICLO))
        self._reprocessar(monkeypatch, ciclo, avaliacao, reimporta=True)
        monkeypatch.setattr(backup, "restaurar_arquivo", lambda *_a, **_k: False)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            codigo = cli._run_pipeline_mensal(9, 2026)

        assert codigo == backup.CODIGO_SAIDA_RESTAURACAO_INCOMPLETA
        assert "RESTAURAÇÃO INCOMPLETA" in caplog.text
        assert "Nada novo para registrar" not in caplog.text

    def test_versionamento_novo_no_reprocessamento_segue_o_ciclo(self, monkeypatch, ciclo):
        avaliacao = AvaliacaoMatriz(
            linhas_rodape=list(self.RODAPE_DO_CICLO),
            alteracoes={"DOC-POP-9": {"v_antiga": "1.0", "v_nova": "2.0"}},
        )
        chamadas = self._reprocessar(monkeypatch, ciclo, avaliacao)

        assert cli._run_pipeline_mensal(9, 2026) == 0

        assert chamadas == ["atualizar"]
        assert _ler(ciclo.destino) == "setembro DEPOIS do ciclo"

    @pytest.mark.parametrize(
        "rodape",
        [[], ["Referente Agosto 2026", "Não houveram alterações nos POPs"]],
        ids=["sem rodape", "so rodape de mes anterior"],
    )
    def test_primeira_execucao_sem_versionamento_escreve_o_rodape(
        self, monkeypatch, ciclo, rodape
    ):
        avaliacao = AvaliacaoMatriz(linhas_rodape=rodape)
        chamadas = self._reprocessar(monkeypatch, ciclo, avaliacao)

        assert cli._run_pipeline_mensal(9, 2026) == 0

        assert chamadas == ["atualizar"]
