from types import SimpleNamespace

from treinamentos_its.reprocessamento import (
    _abas_com_titulo_de_rodape,
    motivos_reprocessamento,
)


MESES = {
    "atual": "Agosto",
    "ano_atual": 2026,
    "anterior": "Julho",
    "ano_ant_relativo": 2026,
}


def test_motivos_usam_so_planos_e_rodape():
    motivos = motivos_reprocessamento(
        MESES,
        planos_macs_do_mes_existe=True,
        abas_com_rodape_do_mes=["Gerente"],
    )
    assert len(motivos) == 2
    assert "Planos e Macs Agosto 2026" in motivos[0]
    assert "Referente Agosto 2026" in motivos[1]  # mês do ciclo, não o de origem


def test_sem_sinal_nao_e_reprocessamento():
    assert motivos_reprocessamento(
        MESES, planos_macs_do_mes_existe=False
    ) == []


class _Range:
    def __init__(self, value):
        self.value = value


class _Sheet:
    def __init__(self, name, values):
        self.name = name
        self.values = values
        self.used_range = SimpleNamespace(last_cell=SimpleNamespace(row=len(values)))

    def range(self, _inicio, _fim):
        return _Range(self.values)


class _Sheets(list):
    def __getitem__(self, key):
        if isinstance(key, str):
            return next(sheet for sheet in self if sheet.name == key)
        return super().__getitem__(key)


def test_acha_rodape_so_nas_abas_pedidas():
    wb = SimpleNamespace(
        sheets=_Sheets(
            [
                _Sheet("Gerente", ["x", " Referente Julho 2026 "]),
                _Sheet("Resumo", ["Referente Julho 2026"]),
            ]
        )
    )
    assert _abas_com_titulo_de_rodape(
        wb, "Referente Julho 2026", ["Gerente"]
    ) == ["Gerente"]
