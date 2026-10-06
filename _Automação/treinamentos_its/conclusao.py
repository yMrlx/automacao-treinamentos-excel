"""Registro transacional de conclusão no Planos e Macs."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from treinamentos_its.backup import (
    CODIGO_SAIDA_RESTAURACAO_INCOMPLETA,
    ErroAoPrepararRestauracao,
    TransacaoArquivo,
    arquivos_em_uso,
    avisar_arquivos_em_uso,
    desligar_autosave,
    instrucoes_de_restauracao_manual,
)
from treinamentos_its.config import (
    ABAS_DE_CARGO,
    MESES,
    PASTA_PLANOS_MACS,
    caminho_copia_pre_colagem,
)
from treinamentos_its.excel_app import abrir_livro, app_excel
from treinamentos_its.dominio import normalizar_status
from treinamentos_its.excel_utils import limpar_celula_status, normalize_name
from treinamentos_its.logger import log
from treinamentos_its.planos_macs import mapear_codigos_da_aba


@dataclass(frozen=True)
class OcorrenciaPlanos:
    aba: str
    linha: int
    coluna_planejado: int
    linha_codigo: int
    status: object = None


class PlanilhaMudouDuranteAConclusao(RuntimeError):
    """A coordenada localizada deixou de representar a mesma pessoa+curso."""


def _localizar_na_aba(
    ws, nome_norm: str, codigo_norm: str
) -> list[OcorrenciaPlanos]:
    blocos = mapear_codigos_da_aba(ws)
    ocorrencias: list[OcorrenciaPlanos] = []
    for bloco in blocos:
        colunas = [
            coluna
            for coluna, codigo in bloco["codigos"].items()
            if str(codigo).strip() == codigo_norm
        ]
        if not colunas:
            continue
        # O fim vem pronto do `planos_macs._ler_blocos` (próximo cabeçalho
        # "Treinamentos", com ou sem código) — mesmo fim que o ciclo usa.
        for linha in range(bloco["linha_func"], bloco["linha_fim"]):
            nome = ws.range(linha, 2).value
            if not nome or normalize_name(nome) != nome_norm:
                continue
            for coluna in colunas:
                ocorrencias.append(
                    OcorrenciaPlanos(
                        ws.name, linha, coluna, bloco["linha_cod"],
                        ws.range(linha, coluna + 1).value,
                    )
                )
    return ocorrencias


def _localizar_ocorrencias(
    arquivo: Path, nome_norm: str, codigo_norm: str
) -> list[OcorrenciaPlanos]:
    with app_excel() as app:
        wb = abrir_livro(app, arquivo, read_only=True)
        try:
            nomes = {ws.name for ws in wb.sheets}
            ocorrencias: list[OcorrenciaPlanos] = []
            for nome_aba in ABAS_DE_CARGO:
                if nome_aba in nomes:
                    ocorrencias.extend(
                        _localizar_na_aba(wb.sheets[nome_aba], nome_norm, codigo_norm)
                    )
            return ocorrencias
        finally:
            wb.close()


def _identidade_confere(
    ws, ocorrencia: OcorrenciaPlanos, nome_norm: str, codigo_norm: str
) -> bool:
    nome = ws.range(ocorrencia.linha, 2).value
    codigo = ws.range(
        ocorrencia.linha_codigo, ocorrencia.coluna_planejado
    ).value
    return (
        bool(nome)
        and normalize_name(nome) == nome_norm
        and str(codigo or "").strip() == codigo_norm
    )


def _marcar_ok(ws, ocorrencia: OcorrenciaPlanos) -> None:
    planejado = ws.range(ocorrencia.linha, ocorrencia.coluna_planejado)
    executado = ws.range(ocorrencia.linha, ocorrencia.coluna_planejado + 1)
    limpar_celula_status(planejado)
    executado.value = "OK"
    executado.color = None
    executado.font.color = (0, 0, 0)


def _status_pendente(
    arquivo: Path, ocorrencia: OcorrenciaPlanos, status: object,
    nome_norm: str, codigo_norm: str,
) -> bool:
    status_norm = normalizar_status(status)
    if status_norm in {"ON TIME", "ATRASADO"} or _esta_vazio(status):
        return True
    local = (
        f"'{arquivo.name}' | '{ocorrencia.aba}' L{ocorrencia.linha} "
        f"C{ocorrencia.coluna_planejado + 1} (Executado) | "
        f"'{nome_norm}' + '{codigo_norm}'"
    )
    if status_norm == "OK":
        log.info("%s: já estava OK; Planejado e Executado preservados.", local)
    else:
        motivo = (
            "status protegido"
            if status_norm in {"HSE", "OUTROS", "GLOBAL", "NA", "N/A"}
            else "status desconhecido"
        )
        log.warning(
            "%s: %s (%r); não é pendente, Planejado e Executado NÃO foram mexidos.",
            local, motivo, status,
        )
    return False


def _salvar_conclusao(
    arquivo: Path,
    ocorrencias: list[OcorrenciaPlanos],
    nome_norm: str,
    codigo_norm: str,
) -> list[OcorrenciaPlanos]:
    """Revalida identidade e status; devolve só as ocorrências escritas."""
    with app_excel() as app:
        wb = abrir_livro(app, arquivo)
        try:
            desligar_autosave(
                wb, descricao=f"Planos e Macs '{arquivo.name}' durante a conclusão"
            )
            nomes = {ws.name for ws in wb.sheets}
            divergencias: list[str] = []
            for ocorrencia in ocorrencias:
                if ocorrencia.aba not in nomes:
                    divergencias.append(f"a aba '{ocorrencia.aba}' não existe mais")
                    continue
                ws = wb.sheets[ocorrencia.aba]
                if not _identidade_confere(ws, ocorrencia, nome_norm, codigo_norm):
                    divergencias.append(
                        f"{ocorrencia.aba} L{ocorrencia.linha} C{ocorrencia.coluna_planejado} mudou"
                    )
            if divergencias:
                raise PlanilhaMudouDuranteAConclusao(
                    "O Planos e Macs mudou entre a leitura e a escrita: "
                    + "; ".join(divergencias)
                )
            pendentes = [
                ocorrencia for ocorrencia in ocorrencias
                if _status_pendente(
                    arquivo, ocorrencia,
                    wb.sheets[ocorrencia.aba].range(
                        ocorrencia.linha, ocorrencia.coluna_planejado + 1
                    ).value,
                    nome_norm, codigo_norm,
                )
            ]
            for ocorrencia in pendentes:
                _marcar_ok(wb.sheets[ocorrencia.aba], ocorrencia)
            if pendentes:
                wb.save()
            return pendentes
        finally:
            wb.close()


def _esta_vazio(valor: object) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _reabrir_e_conferir(
    arquivo: Path,
    ocorrencias: list[OcorrenciaPlanos],
    nome_norm: str,
    codigo_norm: str,
) -> bool:
    with app_excel() as app:
        wb = abrir_livro(app, arquivo, read_only=True)
        try:
            nomes = {ws.name for ws in wb.sheets}
            for ocorrencia in ocorrencias:
                if ocorrencia.aba not in nomes:
                    return False
                ws = wb.sheets[ocorrencia.aba]
                if not _identidade_confere(ws, ocorrencia, nome_norm, codigo_norm):
                    return False
                planejado = ws.range(
                    ocorrencia.linha, ocorrencia.coluna_planejado
                ).value
                executado = ws.range(
                    ocorrencia.linha, ocorrencia.coluna_planejado + 1
                ).value
                if not _esta_vazio(planejado) or normalizar_status(executado) != "OK":
                    return False
        finally:
            wb.close()
    return True


def _mes_ano_do_planos_macs(
    data_conclusao: datetime,
    mes_planos: int | None,
    ano_planos: int | None,
) -> tuple[int, int] | None:
    if (mes_planos is None) != (ano_planos is None):
        log.error("Informe --mes e --ano juntos (ou nenhum dos dois).")
        return None
    if mes_planos is not None:
        if not 1 <= mes_planos <= 12:
            log.error("--mes deve estar entre 1 e 12.")
            return None
        return mes_planos, int(ano_planos)
    return data_conclusao.month, data_conclusao.year


def registrar_conclusao(
    nome: str,
    codigo_curso: str,
    data_conclusao: datetime | None = None,
    mes_planos: int | None = None,
    ano_planos: int | None = None,
) -> int:
    """Grava ``OK`` e limpa Planejado apenas nas ocorrências pendentes."""
    data_conclusao = data_conclusao or datetime.now()
    alvo = _mes_ano_do_planos_macs(data_conclusao, mes_planos, ano_planos)
    if alvo is None:
        return 1
    mes, ano = alvo
    arquivo = PASTA_PLANOS_MACS / f"Planos e Macs {MESES[mes - 1]} {ano}.xlsx"
    if not arquivo.exists():
        log.error("Planos e Macs do mês não encontrado: %s", arquivo.resolve())
        return 1

    nome_norm = normalize_name(nome)
    codigo_norm = str(codigo_curso or "").strip()
    try:
        ocorrencias = _localizar_ocorrencias(arquivo, nome_norm, codigo_norm)
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao localizar a conclusão: %s", erro)
        return 1
    if not ocorrencias:
        log.error(
            "Não encontrei '%s' + curso '%s' nas abas de cargo de %s.",
            nome,
            codigo_norm,
            arquivo.name,
        )
        return 1

    pendentes = [
        ocorrencia for ocorrencia in ocorrencias
        if _status_pendente(
            arquivo, ocorrencia, ocorrencia.status, nome_norm, codigo_norm
        )
    ]
    if not pendentes:
        log.info("Nada foi alterado: '%s' + '%s' sem ocorrência pendente em %s.",
                 nome, codigo_norm, arquivo.name)
        return 0

    em_uso = arquivos_em_uso([arquivo])
    if em_uso:
        avisar_arquivos_em_uso(em_uso)
        return 1

    transacao = TransacaoArquivo(
        arquivo,
        caminho_copia_pre_colagem(arquivo.name),
        descricao=f"Planos e Macs '{arquivo.name}' durante a conclusão",
    )
    sucesso = False
    gravadas: list[OcorrenciaPlanos] = []
    try:
        with transacao:
            gravadas = _salvar_conclusao(arquivo, pendentes, nome_norm, codigo_norm)
            if not gravadas:
                # Ninguém foi alterado: não restaurar sobre uma edição externa
                # que tornou as células protegidas entre a leitura e a escrita.
                transacao.confirmar()
                sucesso = True
            elif not _reabrir_e_conferir(
                arquivo, gravadas, nome_norm, codigo_norm
            ):
                log.error(
                    "A conclusão não permaneceu salva depois de reabrir; o arquivo será restaurado."
                )
            else:
                transacao.confirmar()
                sucesso = True
    except ErroAoPrepararRestauracao as erro:
        log.error("%s", erro)
    except PlanilhaMudouDuranteAConclusao as erro:
        log.error("%s", erro)
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao registrar a conclusão: %s", erro)

    if sucesso:
        if not gravadas:
            log.info("Nada foi alterado: as ocorrências deixaram de ser pendentes antes da escrita.")
            return 0
        log.info(
            "Conclusão confirmada após reabrir: '%s' + '%s' em %s (%d ocorrência(s)); "
            "status OK e Planejado vazio.",
            nome,
            codigo_norm,
            arquivo.name,
            len(gravadas),
        )
        return 0
    if transacao.falhas_de_restauracao:
        log.error(
            "RESTAURAÇÃO INCOMPLETA: a conclusão falhou e o arquivo não voltou ao estado anterior:"
        )
        for instrucao in instrucoes_de_restauracao_manual(
            transacao.falhas_de_restauracao
        ):
            log.error("  - %s", instrucao)
        return CODIGO_SAIDA_RESTAURACAO_INCOMPLETA
    return 1
