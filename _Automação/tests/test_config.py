from datetime import datetime

from treinamentos_its import config


def test_config_tem_apenas_planos_como_fonte_de_dados():
    assert config.PASTA_PLANOS_MACS == config.PATH_BASE / "Planos e Macs"
    assert not hasattr(config, "ARQUIVO_PRINCIPAL")
    assert not hasattr(config, "PASTA_PROCEDIMENTOS")


def test_abas_de_cargo_sao_lista_unica_sem_mapeamento():
    assert len(config.ABAS_DE_CARGO) == 6
    assert len(set(config.ABAS_DE_CARGO)) == len(config.ABAS_DE_CARGO)
    assert not hasattr(config, "MAP_ABAS")


def test_caminho_copia_pre_colagem_tem_timestamp(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "PASTA_PRE_COLAGEM", tmp_path)
    caminho = config.caminho_copia_pre_colagem(
        "Planos e Macs Setembro 2026.xlsx", datetime(2026, 9, 25, 10, 11, 12)
    )
    assert caminho == tmp_path / (
        "Planos e Macs Setembro 2026 — antes de 2026-09-25 10-11-12.xlsx"
    )
