"""Guarda contra reprocessar sem querer o Planos e Macs mensal."""

from __future__ import annotations

from treinamentos_its.logger import log


def motivos_reprocessamento(
    meses: dict,
    *,
    planos_macs_do_mes_existe: bool,
    abas_com_rodape_do_mes: list[str] | None = None,
) -> list[str]:
    motivos: list[str] = []
    if planos_macs_do_mes_existe:
        motivos.append(
            f"a planilha 'Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx' "
            "de destino já existe"
        )
    abas = abas_com_rodape_do_mes or []
    if abas:
        motivos.append(
            f"o rodapé 'Referente {meses['atual']} {meses['ano_atual']}' "
            f"já está escrito em: {', '.join(abas)}"
        )
    return motivos


def _abas_com_titulo_de_rodape(wb, titulo: str, abas_de_cargo=None) -> list[str]:
    nomes = {ws.name for ws in wb.sheets}
    achou: list[str] = []
    for nome_aba in abas_de_cargo or nomes:
        if nome_aba not in nomes:
            continue
        ws = wb.sheets[nome_aba]
        try:
            ultima = ws.used_range.last_cell.row
            col_a = ws.range((1, 1), (ultima, 1)).value
        except Exception as erro:  # noqa: BLE001
            log.warning("Não consegui ler a aba '%s' do Planos e Macs: %s", ws.name, erro)
            continue
        if not isinstance(col_a, list):
            col_a = [col_a]
        if col_a and isinstance(col_a[0], list):
            col_a = [linha[0] if linha else None for linha in col_a]
        if any(isinstance(v, str) and v.strip() == titulo for v in col_a):
            achou.append(ws.name)
    return achou


def sinais_de_reprocessamento(meses: dict) -> list[str]:
    from treinamentos_its.config import ABAS_DE_CARGO, PASTA_PLANOS_MACS
    from treinamentos_its.excel_app import abrir_livro, app_excel

    destino = PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    )
    titulo = f"Referente {meses['atual']} {meses['ano_atual']}"
    abas: list[str] = []
    if destino.exists():
        try:
            with app_excel() as app:
                wb = abrir_livro(app, destino, read_only=True)
                try:
                    abas = _abas_com_titulo_de_rodape(wb, titulo, ABAS_DE_CARGO)
                finally:
                    wb.close()
        except Exception as erro:  # noqa: BLE001
            log.warning(
                "Não consegui checar os rodapés de '%s' (%s); seguindo com o sinal "
                "de que o arquivo já existe.",
                destino.name,
                erro,
            )
    return motivos_reprocessamento(
        meses,
        planos_macs_do_mes_existe=destino.exists(),
        abas_com_rodape_do_mes=abas,
    )
