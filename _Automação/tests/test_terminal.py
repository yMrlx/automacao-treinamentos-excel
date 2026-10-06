import builtins

from treinamentos_its.terminal import perguntar_sim_nao


def _responder(monkeypatch, resposta):
    monkeypatch.setattr(builtins, "input", lambda _pergunta="": resposta)


def test_perguntar_sim_nao_aceita_s(monkeypatch):
    _responder(monkeypatch, "s")
    assert perguntar_sim_nao("? ") is True


def test_perguntar_sim_nao_aceita_sim_com_espacos_e_maiuscula(monkeypatch):
    _responder(monkeypatch, "  SIM  ")
    assert perguntar_sim_nao("? ") is True


def test_perguntar_sim_nao_recusa_n(monkeypatch):
    _responder(monkeypatch, "n")
    assert perguntar_sim_nao("? ") is False


def test_perguntar_sim_nao_qualquer_outra_coisa_e_nao(monkeypatch):
    _responder(monkeypatch, "talvez")
    assert perguntar_sim_nao("? ") is False


def test_perguntar_sim_nao_eof_assume_nao_por_padrao(monkeypatch):
    def _estoura(_pergunta=""):
        raise EOFError

    monkeypatch.setattr(builtins, "input", _estoura)
    assert perguntar_sim_nao("? ") is False


def test_perguntar_sim_nao_eof_pode_assumir_sim(monkeypatch):
    def _estoura(_pergunta=""):
        raise EOFError

    monkeypatch.setattr(builtins, "input", _estoura)
    assert perguntar_sim_nao("? ", assumir_em_eof=True) is True
