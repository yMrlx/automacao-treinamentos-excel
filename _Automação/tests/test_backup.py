from pathlib import Path

from treinamentos_its import backup


def test_transacao_restaura_arquivo_existente(tmp_path):
    arquivo = tmp_path / "planos.xlsx"
    copia = tmp_path / "copias" / "antes.xlsx"
    arquivo.write_text("antes", encoding="utf-8")

    with backup.TransacaoArquivo(arquivo, copia, descricao="Planos"):
        arquivo.write_text("durante", encoding="utf-8")

    assert arquivo.read_text(encoding="utf-8") == "antes"


def test_transacao_confirmada_mantem_resultado(tmp_path):
    arquivo = tmp_path / "planos.xlsx"
    arquivo.write_text("antes", encoding="utf-8")
    with backup.TransacaoArquivo(
        arquivo, tmp_path / "antes.xlsx", descricao="Planos"
    ) as transacao:
        arquivo.write_text("depois", encoding="utf-8")
        transacao.confirmar()
    assert arquivo.read_text(encoding="utf-8") == "depois"


def test_transacao_remove_arquivo_que_nasceu_na_operacao(tmp_path):
    arquivo = tmp_path / "novo.xlsx"
    with backup.TransacaoArquivo(
        arquivo,
        tmp_path / "nao-usada.xlsx",
        descricao="Planos novo",
        exigir_existente=False,
    ):
        arquivo.write_text("incompleto", encoding="utf-8")
    assert not arquivo.exists()


def test_copia_de_restauracao_nao_sobrescreve_nome_repetido(tmp_path):
    arquivo = tmp_path / "planos.xlsx"
    destino = tmp_path / "antes.xlsx"
    arquivo.write_text("x", encoding="utf-8")
    destino.write_text("primeira", encoding="utf-8")

    criada = backup.copiar_para_restauracao(arquivo, destino, descricao="Planos")

    assert criada == tmp_path / "antes (2).xlsx"
    assert destino.read_text(encoding="utf-8") == "primeira"


def test_instrucao_manual_distingue_restaurar_e_apagar(tmp_path):
    existente = backup.TransacaoArquivo(
        tmp_path / "a.xlsx", tmp_path / "a-antes.xlsx", descricao="existente"
    )
    existente.copia = Path("copia.xlsx")
    novo = backup.TransacaoArquivo(
        tmp_path / "b.xlsx",
        tmp_path / "b-antes.xlsx",
        descricao="novo",
        exigir_existente=False,
    )
    assert "copie" in backup.instrucoes_de_restauracao_manual([existente])[0]
    assert "apague" in backup.instrucoes_de_restauracao_manual([novo])[0]


# =============================================================================
# Cobertura recriada depois do a5dad36: restauração que falha, os textos que o
# painel/operadora leem e a transação composta. Só arquivos em tmp_path.
# =============================================================================

import logging
from types import SimpleNamespace

import pytest


def _copia_quebrada(*_args, **_kwargs):
    raise PermissionError("arquivo em uso")


class TestRestauracaoQueFalha:
    def test_copia_que_falha_na_volta_e_registrada(self, monkeypatch, tmp_path, caplog):
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("antes", encoding="utf-8")
        copia = tmp_path / "copias" / "antes.xlsx"

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            with backup.TransacaoArquivo(alvo, copia, descricao="Planos") as transacao:
                alvo.write_text("pela metade", encoding="utf-8")
                # A cópia de ida funcionou; a de volta (restauração) vai falhar.
                monkeypatch.setattr(backup.shutil, "copy2", _copia_quebrada)

        assert transacao.restauracao_falhou is True
        assert transacao.falhas_de_restauracao == [transacao]
        assert alvo.read_text(encoding="utf-8") == "pela metade"
        assert copia.read_text(encoding="utf-8") == "antes"  # a cópia ficou
        assert backup.instrucoes_de_restauracao_manual(transacao.falhas_de_restauracao) == [
            f"'{alvo}': com o Excel FECHADO, copie '{copia}' por cima dele"
        ]

    def test_arquivo_novo_que_nao_apaga_e_registrado(self, monkeypatch, tmp_path, caplog):
        alvo = tmp_path / "novo.xlsx"
        unlink_original = Path.unlink

        def _unlink_travado(self, *args, **kwargs):
            if self == alvo:
                raise PermissionError("aberto no Excel")
            return unlink_original(self, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", _unlink_travado)

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            with backup.TransacaoArquivo(
                alvo, tmp_path / "nao-usada.xlsx", descricao="novo", exigir_existente=False
            ) as transacao:
                alvo.write_text("parcial", encoding="utf-8")

        assert transacao.restauracao_falhou is True
        assert transacao.copia is None
        assert "NÃO consegui remover o arquivo incompleto" in caplog.text
        assert backup.instrucoes_de_restauracao_manual(transacao.falhas_de_restauracao) == [
            f"'{alvo}': nasceu nesta execução e ficou pela metade - apague-o à mão"
        ]

    def test_restauracao_que_da_certo_nao_registra_falha(self, tmp_path):
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("antes", encoding="utf-8")
        with backup.TransacaoArquivo(alvo, tmp_path / "c" / "a.xlsx", descricao="x") as transacao:
            alvo.write_text("depois", encoding="utf-8")
        assert alvo.read_text(encoding="utf-8") == "antes"
        assert transacao.restauracao_falhou is False
        assert transacao.falhas_de_restauracao == []

    def test_excecao_dentro_da_transacao_restaura_e_propaga(self, tmp_path):
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("antes", encoding="utf-8")
        with pytest.raises(RuntimeError, match="Excel caiu"):
            with backup.TransacaoArquivo(alvo, tmp_path / "c" / "a.xlsx", descricao="x"):
                alvo.write_text("AutoSave parcial", encoding="utf-8")
                raise RuntimeError("Excel caiu")
        assert alvo.read_text(encoding="utf-8") == "antes"

    def test_sem_conseguir_copiar_na_entrada_nada_comeca(self, monkeypatch, tmp_path, caplog):
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("antes", encoding="utf-8")
        monkeypatch.setattr(backup.shutil, "copy2", _copia_quebrada)
        executou = []
        with pytest.raises(backup.ErroAoPrepararRestauracao, match="Sem cópia de restauração"):
            with backup.TransacaoArquivo(alvo, tmp_path / "c" / "a.xlsx", descricao="Planos"):
                executou.append(True)
        assert executou == []
        assert alvo.read_text(encoding="utf-8") == "antes"

    def test_arquivo_exigido_que_nao_existe(self, tmp_path):
        with pytest.raises(backup.ErroAoPrepararRestauracao, match="não encontrado"):
            with backup.TransacaoArquivo(
                tmp_path / "sumiu.xlsx", tmp_path / "c.xlsx", descricao="Planos"
            ):
                pass

    def test_codigo_de_saida_da_restauracao_incompleta_e_3(self):
        # Os .bat e o painel tratam o 3 como "restaurar à mão".
        assert backup.CODIGO_SAIDA_RESTAURACAO_INCOMPLETA == 3


class TestTextosQueOPainelLe:
    def test_restaurado(self, tmp_path, caplog):
        copia = tmp_path / "copia.xlsx"
        copia.write_text("antes", encoding="utf-8")
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("depois", encoding="utf-8")

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert backup.restaurar_arquivo(copia, alvo, descricao="Planos X") is True

        assert alvo.read_text(encoding="utf-8") == "antes"
        assert f"Planos X foi RESTAURADO a partir de '{copia}'" in caplog.text
        assert caplog.records[-1].levelno == logging.WARNING

    def test_restauracao_falhou(self, monkeypatch, tmp_path, caplog):
        monkeypatch.setattr(backup.shutil, "copy2", _copia_quebrada)
        copia = tmp_path / "copia.xlsx"
        alvo = tmp_path / "alvo.xlsx"

        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert backup.restaurar_arquivo(copia, alvo, descricao="Planos X") is False

        assert "RESTAURAÇÃO FALHOU para Planos X" in caplog.text
        assert f"a partir de '{copia}'" in caplog.text
        assert "NÃO rode o ciclo novamente" in caplog.text
        assert caplog.records[-1].levelno == logging.ERROR

    def test_copia_de_restauracao_anuncia_onde_ficou(self, tmp_path, caplog):
        # "Cópia de restauração de" é o marcador da 1ª etapa do sync/concluir no painel.
        alvo = tmp_path / "alvo.xlsx"
        alvo.write_text("x", encoding="utf-8")
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            criada = backup.copiar_para_restauracao(
                alvo, tmp_path / "c" / "antes.xlsx", descricao="Planos X"
            )
        assert criada.read_text(encoding="utf-8") == "x"
        assert f"Cópia de restauração de Planos X: {criada.resolve()}" in caplog.text


def _par(tmp_path, nome, conteudo=None, exigir=True):
    arquivo = tmp_path / f"{nome}.xlsx"
    if conteudo is not None:
        arquivo.write_text(conteudo, encoding="utf-8")
    transacao = backup.TransacaoArquivo(
        arquivo, tmp_path / "copias" / f"{nome}.xlsx", descricao=nome, exigir_existente=exigir
    )
    return arquivo, transacao


class TestTransacaoArquivos:
    def test_desfaz_todos_os_arquivos(self, tmp_path):
        primeiro, t1 = _par(tmp_path, "primeiro", "antes 1")
        segundo, t2 = _par(tmp_path, "segundo", "antes 2")
        with backup.TransacaoArquivos(t1, t2) as transacao:
            primeiro.write_text("depois 1", encoding="utf-8")
            segundo.write_text("depois 2", encoding="utf-8")
        assert primeiro.read_text(encoding="utf-8") == "antes 1"
        assert segundo.read_text(encoding="utf-8") == "antes 2"
        assert transacao.restauracao_falhou is False
        assert transacao.confirmada is False

    def test_apaga_o_arquivo_que_nasceu_e_restaura_o_que_existia(self, tmp_path):
        existente, t1 = _par(tmp_path, "existente", "antes")
        novo, t2 = _par(tmp_path, "novo", exigir=False)
        with backup.TransacaoArquivos(t1, t2):
            existente.write_text("depois", encoding="utf-8")
            novo.write_text("parcial", encoding="utf-8")
        assert existente.read_text(encoding="utf-8") == "antes"
        assert not novo.exists()

    def test_confirmada_mantem_o_conjunto(self, tmp_path):
        primeiro, t1 = _par(tmp_path, "primeiro", "antes 1")
        segundo, t2 = _par(tmp_path, "segundo", "antes 2")
        with backup.TransacaoArquivos(t1, t2) as transacao:
            primeiro.write_text("depois 1", encoding="utf-8")
            segundo.write_text("depois 2", encoding="utf-8")
            transacao.confirmar()
        assert primeiro.read_text(encoding="utf-8") == "depois 1"
        assert segundo.read_text(encoding="utf-8") == "depois 2"
        assert transacao.confirmada is True and t1.confirmada and t2.confirmada

    def test_diz_quais_arquivos_nao_voltaram(self, monkeypatch, tmp_path):
        primeiro, t1 = _par(tmp_path, "primeiro", "antes 1")
        segundo, t2 = _par(tmp_path, "segundo", "antes 2")
        restaurar_de_verdade = backup.restaurar_arquivo

        def _restaurar(copia, arquivo, *, descricao):
            if Path(arquivo) == segundo:
                return False
            return restaurar_de_verdade(copia, arquivo, descricao=descricao)

        monkeypatch.setattr(backup, "restaurar_arquivo", _restaurar)
        with backup.TransacaoArquivos(t1, t2) as transacao:
            primeiro.write_text("depois 1", encoding="utf-8")
            segundo.write_text("depois 2", encoding="utf-8")

        assert primeiro.read_text(encoding="utf-8") == "antes 1"
        assert transacao.restauracao_falhou is True
        assert [falha.arquivo for falha in transacao.falhas_de_restauracao] == [segundo]
        assert backup.instrucoes_de_restauracao_manual(transacao.falhas_de_restauracao) == [
            f"'{segundo}': com o Excel FECHADO, copie '{t2.copia}' por cima dele"
        ]

    def test_excecao_dentro_desfaz_e_propaga(self, tmp_path):
        primeiro, t1 = _par(tmp_path, "primeiro", "antes 1")
        with pytest.raises(RuntimeError):
            with backup.TransacaoArquivos(t1):
                primeiro.write_text("depois 1", encoding="utf-8")
                raise RuntimeError("Excel caiu")
        assert primeiro.read_text(encoding="utf-8") == "antes 1"

    def test_segunda_que_nao_prepara_desfaz_a_primeira_e_nada_roda(self, tmp_path):
        primeiro, t1 = _par(tmp_path, "primeiro", "antes 1")
        _sumiu, t2 = _par(tmp_path, "sumiu")  # exigido e não existe
        executou = []
        with pytest.raises(backup.ErroAoPrepararRestauracao):
            with backup.TransacaoArquivos(t1, t2):
                executou.append(True)
        assert executou == []
        assert primeiro.read_text(encoding="utf-8") == "antes 1"
        assert t1.confirmada is False

    def test_saida_sem_entrar_nao_quebra(self, tmp_path):
        _a, t1 = _par(tmp_path, "a", "x")
        assert backup.TransacaoArquivos(t1).__exit__(None, None, None) is False


class TestDesligarAutosave:
    def test_desliga_quando_esta_ligado(self, caplog):
        wb = SimpleNamespace(api=SimpleNamespace(AutoSaveOn=True))
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert backup.desligar_autosave(wb, descricao="Planos X") is True
        assert wb.api.AutoSaveOn is False
        assert "AutoSave do OneDrive DESLIGADO em Planos X" in caplog.text

    def test_ja_desligado_nao_mexe(self, caplog):
        wb = SimpleNamespace(api=SimpleNamespace(AutoSaveOn=False))
        with caplog.at_level(logging.INFO, logger="treinamentos_its"):
            assert backup.desligar_autosave(wb, descricao="Planos X") is True
        assert "DESLIGADO" not in caplog.text

    def test_erro_ao_desligar_so_avisa(self, caplog):
        class _Api:
            AutoSaveOn = True

            def __setattr__(self, nome, valor):
                raise RuntimeError("COM recusou")

        wb = SimpleNamespace(api=_Api())
        assert backup.desligar_autosave(wb, descricao="Planos X") is False
        assert "Não consegui desligar o AutoSave em Planos X" in caplog.text
