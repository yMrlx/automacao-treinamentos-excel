from treinamentos_its import ambiente_teste


def _montar_raiz_fake(tmp_path):
    """Cria uma raiz de projeto de mentira só com Planos e Macs."""
    raiz = tmp_path / "raiz"
    (raiz / "Planos e Macs").mkdir(parents=True)
    (raiz / "Planos e Macs" / "Planos e Macs Julho 2026.xlsx").write_text("a")
    (raiz / "Planos e Macs" / "~$lock.xlsx").write_text("lock")
    return raiz


def _redirecionar(monkeypatch, tmp_path):
    raiz = _montar_raiz_fake(tmp_path)
    sandbox = tmp_path / "_ambiente-teste"
    monkeypatch.setattr(ambiente_teste, "RAIZ_PROJETO", raiz)
    monkeypatch.setattr(ambiente_teste, "PASTA_SANDBOX", sandbox)
    return raiz, sandbox


def test_preparar_sandbox_copia_planilhas_na_primeira_vez(monkeypatch, tmp_path):
    _raiz, sandbox = _redirecionar(monkeypatch, tmp_path)

    resultado = ambiente_teste.preparar_sandbox()

    assert resultado == sandbox
    assert (sandbox / "Planos e Macs" / "Planos e Macs Julho 2026.xlsx").exists()
    assert not (sandbox / "Procedimentos").exists()
    # arquivo de lock do Excel não deve ser copiado
    assert not (sandbox / "Planos e Macs" / "~$lock.xlsx").exists()


def test_preparar_sandbox_nao_recopia_quando_ja_tem_planilha(monkeypatch, tmp_path):
    _raiz, sandbox = _redirecionar(monkeypatch, tmp_path)
    ambiente_teste.preparar_sandbox()

    # Simula edição no sandbox; uma 2ª chamada não pode sobrescrever.
    alvo = sandbox / "Planos e Macs" / "Planos e Macs Julho 2026.xlsx"
    alvo.write_text("editado no teste")

    ambiente_teste.preparar_sandbox()

    assert alvo.read_text() == "editado no teste"


def test_resetar_sandbox_recria_do_zero(monkeypatch, tmp_path):
    _raiz, sandbox = _redirecionar(monkeypatch, tmp_path)
    ambiente_teste.preparar_sandbox()
    lixo = sandbox / "Planos e Macs" / "Planos e Macs Agosto 2026.xlsx"
    lixo.write_text("resultado de uma run anterior")

    ambiente_teste.resetar_sandbox()

    assert not lixo.exists()
    assert (sandbox / "Planos e Macs" / "Planos e Macs Julho 2026.xlsx").exists()


# --- "isto é modo teste?" -- a decisão que não pode divergir ------------------
#
# O furo: o `main.py` procurava a string exata "--teste" e o argparse aceitava
# abreviação. `python main.py sync --tes` = sandbox NÃO preparado + argparse
# dizendo que é teste = escrita nas planilhas de PRODUÇÃO achando que é cópia.


def test_modo_teste_pedido_reconhece_a_flag_inteira():
    assert ambiente_teste.modo_teste_pedido(["main.py", "sync", "--teste"]) is True
    assert ambiente_teste.modo_teste_pedido(["main.py", "conferir", "--teste"]) is True
    assert ambiente_teste.modo_teste_pedido(
        ["main.py", "sync", "--mes", "7", "--ano", "2026", "--teste"]
    ) is True


def test_modo_teste_pedido_ignora_abreviacao():
    # A abreviação NÃO prepara o sandbox aqui - e, do outro lado, o argparse
    # (allow_abbrev=False) também não a aceita mais: as duas pontas concordam.
    assert ambiente_teste.modo_teste_pedido(["main.py", "sync", "--tes"]) is False
    assert ambiente_teste.modo_teste_pedido(["main.py", "sync", "--te"]) is False


def test_modo_teste_pedido_so_vale_pros_comandos_que_leem_planilha():
    # "concluir" passou a aceitar --teste em 2026-09-21 (a GUI tem um botao pra
    # ele, e sem a flag esse botao escreveria em producao com a tela dizendo
    # "ambiente de teste"). Os comandos que NAO mexem nas planilhas do mes
    # continuam de fora.
    assert ambiente_teste.modo_teste_pedido(["main.py", "concluir", "--teste"]) is True
    assert ambiente_teste.modo_teste_pedido(["main.py", "resetar-teste", "--teste"]) is False
    assert ambiente_teste.modo_teste_pedido(["main.py", "limpar-rodapes", "--teste"]) is False
    assert ambiente_teste.modo_teste_pedido(["main.py", "--teste"]) is False
    assert ambiente_teste.modo_teste_pedido(["main.py"]) is False
    assert ambiente_teste.modo_teste_pedido([]) is False


def test_modo_teste_ativo_compara_com_o_sandbox(monkeypatch, tmp_path):
    _raiz, sandbox = _redirecionar(monkeypatch, tmp_path)
    sandbox.mkdir(parents=True, exist_ok=True)

    assert ambiente_teste.modo_teste_ativo(sandbox) is True
    assert ambiente_teste.modo_teste_ativo(tmp_path / "raiz") is False
