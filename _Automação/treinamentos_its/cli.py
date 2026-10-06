"""Interface de linha de comando da automação."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from treinamentos_its.backup import (
    CODIGO_SAIDA_RESTAURACAO_INCOMPLETA,
    ErroAoPrepararRestauracao,
    TransacaoArquivo,
    arquivos_em_uso,
    avisar_arquivos_em_uso,
    instrucoes_de_restauracao_manual,
)
from treinamentos_its.conclusao import registrar_conclusao
from treinamentos_its.conferir import conferir_matriz
from treinamentos_its.dominio import rodape_do_ciclo_ja_escrito
from treinamentos_its.logger import log
from treinamentos_its.planos_macs import criar_planos_do_mes, atualizar_planos_e_macs
from treinamentos_its.reprocessamento import sinais_de_reprocessamento
from treinamentos_its.sync import avaliar_arquivo_matriz, obter_meses
from treinamentos_its.terminal import perguntar_sim_nao


def _adicionar_referencia(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--mes", type=int,
        help="Mês dos POPs recebidos = do Planos e Macs que vai ser criado (1 a 12).",
    )
    parser.add_argument(
        "--ano", type=int,
        help="Ano desse mesmo mês dos POPs (ex.: POPs de Dezembro/2026 → --mes 12 --ano 2026).",
    )


def _adicionar_teste(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--teste",
        action="store_true",
        help="Roda contra o sandbox _Automação/_ambiente-teste/.",
    )


def _novo_parser(**kwargs) -> argparse.ArgumentParser:
    return argparse.ArgumentParser(allow_abbrev=False, **kwargs)


def _abortar_se_teste_sem_sandbox(args: argparse.Namespace) -> None:
    if not getattr(args, "teste", False):
        return
    from treinamentos_its import config
    from treinamentos_its.ambiente_teste import PASTA_SANDBOX, modo_teste_ativo

    if modo_teste_ativo(config.PATH_BASE):
        return
    log.error(
        "'--teste' foi pedido, mas a base em uso é '%s' e não o sandbox '%s'. "
        "Nada foi executado.",
        config.PATH_BASE,
        PASTA_SANDBOX,
    )
    sys.exit(1)


def _caminhos_do_ciclo(meses: dict) -> tuple[Path, Path]:
    from treinamentos_its import config

    origem = config.PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['anterior']} {meses['ano_ant_relativo']}.xlsx"
    )
    destino = config.PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    )
    return origem, destino


def _transacao_do_ciclo(meses: dict) -> TransacaoArquivo:
    """Protege somente o Planos e Macs do mês novo."""
    from treinamentos_its import config

    _origem, destino = _caminhos_do_ciclo(meses)
    return TransacaoArquivo(
        destino,
        config.caminho_copia_pre_colagem(destino.name),
        descricao=f"Planos e Macs '{destino.name}'",
        exigir_existente=False,
    )


def _preparar_planos_novo(meses: dict, vigentes=None) -> Path | None:
    origem, destino = _caminhos_do_ciclo(meses)
    if not criar_planos_do_mes(origem, destino):
        return None
    if vigentes:
        from treinamentos_its.vigentes import importar_vigentes

        arquivo_vigentes = Path(vigentes)
        if not arquivo_vigentes.exists():
            log.error("Export de POPs vigentes não encontrado: %s", arquivo_vigentes)
            return None
        if not importar_vigentes(arquivo_vigentes, destino):
            log.error("A importação dos POPs vigentes falhou; o ciclo foi interrompido.")
            return None
        log.info("POPs vigentes importados no arquivo do mês novo: %s", destino.name)
    else:
        log.info(
            "Sem --vigentes: usando a Matriz do arquivo novo como veio da cópia de '%s'.",
            origem.name,
        )
    return destino


# Desfecho interno de `_executar_etapas_do_ciclo` (nunca vira código de saída):
# o mês já foi fechado e não apareceu versionamento novo.
NADA_NOVO = -1


def _executar_etapas_do_ciclo(meses: dict, vigentes=None) -> int:
    destino = _preparar_planos_novo(meses, vigentes)
    if destino is None:
        return 1
    avaliacao = avaliar_arquivo_matriz(destino)
    if avaliacao is None or not avaliacao.pode_executar:
        log.error(
            "A validação (Matriz e abas de cargo) bloqueou o ciclo - motivo(s) acima. "
            "Nada será mantido."
        )
        return 1
    # Decisão de negócio (2026-10-02): reprocessar um mês já fechado sem nenhum
    # versionamento novo não grava nada — antes empilhava um segundo bloco
    # "Não houveram alterações" no rodapé, contradizendo o primeiro.
    titulo = f"Referente {meses['atual']} {meses['ano_atual']}"
    if not avaliacao.alteracoes and rodape_do_ciclo_ja_escrito(
        avaliacao.linhas_rodape, titulo
    ):
        return NADA_NOVO
    return 0 if atualizar_planos_e_macs(avaliacao, meses, destino) else 1


def _relatar_desfazimento_do_ciclo(transacao: TransacaoArquivo) -> int:
    falhas = transacao.falhas_de_restauracao
    if not falhas:
        log.error(
            "O ciclo NÃO terminou; o Planos e Macs do mês novo voltou ao estado anterior."
        )
        return 1
    log.error(
        "RESTAURAÇÃO INCOMPLETA: o ciclo falhou e o Planos e Macs não voltou ao "
        "estado anterior. NÃO rode de novo antes de resolver à mão:"
    )
    for instrucao in instrucoes_de_restauracao_manual(falhas):
        log.error("  - %s", instrucao)
    return CODIGO_SAIDA_RESTAURACAO_INCOMPLETA


def _run_pipeline_mensal(
    mes: int | None = None, ano: int | None = None, vigentes=None
) -> int:
    try:
        meses = obter_meses(mes, ano)
    except ValueError as erro:
        log.error("Referência inválida: %s", erro)
        return 1

    log.info("%s", "=" * 55)
    log.info("Automação de Treinamentos - Time de TI")
    motivos = sinais_de_reprocessamento(meses)
    if motivos:
        log.warning(
            "O ciclo de %s/%s parece JÁ ter sido processado:",
            meses["atual"],
            meses["ano_atual"],
        )
        for motivo in motivos:
            log.warning("  - %s", motivo)
        if not perguntar_sim_nao("Rodar mesmo assim? (s/n): "):
            log.info("Ok, não vou reprocessar. Nada foi alterado.")
            return 0

    em_uso = arquivos_em_uso(_caminhos_do_ciclo(meses))
    if em_uso:
        avisar_arquivos_em_uso(em_uso)
        return 1

    transacao = _transacao_do_ciclo(meses)
    protegido = False
    codigo = 1
    try:
        with transacao:
            protegido = True
            codigo = _executar_etapas_do_ciclo(meses, vigentes)
            if codigo == 0:
                transacao.confirmar()
    except ErroAoPrepararRestauracao as erro:
        log.error("%s", erro)
        return 1
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro inesperado no ciclo: %s", erro)
        codigo = 1
    if codigo == 0:
        return 0
    if codigo == NADA_NOVO:
        # A transação não foi confirmada: o Planos do mês novo voltou a ser
        # exatamente o de antes (inclusive uma reimportação de vigentes).
        if transacao.falhas_de_restauracao:
            return _relatar_desfazimento_do_ciclo(transacao)
        # "Nada foi alterado" é o marcador que o painel lê pra mostrar desfecho
        # neutro (plano_execucao.MARCADOR_NADA_ALTERADO).
        log.info(
            "Nada novo para registrar: o rodapé 'Referente %s %s' já existe e não "
            "apareceu nenhum versionamento novo. Nada foi alterado no Planos e Macs "
            "de %s (abas de cargo, rodapé e Matriz ficaram como estavam).",
            meses["atual"],
            meses["ano_atual"],
            meses["atual"],
        )
        return 0
    return _relatar_desfazimento_do_ciclo(transacao) if protegido else codigo


def construir_parser() -> argparse.ArgumentParser:
    parser = _novo_parser(description="Automação de Treinamentos - Time de TI")
    subparsers = parser.add_subparsers(dest="comando", parser_class=_novo_parser)

    sync_parser = subparsers.add_parser("sync", help="Fecha o ciclo mensal no Planos e Macs.")
    _adicionar_referencia(sync_parser)
    _adicionar_teste(sync_parser)
    sync_parser.add_argument(
        "--vigentes",
        help="Export 'POPs aprovados e efetivos' a importar no arquivo do mês novo.",
    )

    concluir_parser = subparsers.add_parser("concluir", help="Marca um curso como concluído.")
    concluir_parser.add_argument("nome")
    concluir_parser.add_argument("codigo")
    concluir_parser.add_argument("--data", help="Data de conclusão em dd/mm/aaaa.")
    concluir_parser.add_argument("--mes", type=int)
    concluir_parser.add_argument("--ano", type=int)
    _adicionar_teste(concluir_parser)

    conferir_parser = subparsers.add_parser(
        "conferir", help="Checagem sem alteração do ciclo mensal."
    )
    _adicionar_referencia(conferir_parser)
    _adicionar_teste(conferir_parser)
    conferir_parser.add_argument("--pops-vigentes")
    conferir_parser.add_argument("--pops-obsoletos")

    subparsers.add_parser("resetar-teste", help="Recria o sandbox de teste.")
    limpar_parser = subparsers.add_parser(
        "limpar-rodapes", help="Dry-run da ferramenta congelada de rodapés."
    )
    limpar_parser.add_argument("arquivo")
    limpar_parser.add_argument("--aplicar", action="store_true")
    return parser


def main() -> None:
    parser = construir_parser()
    args = parser.parse_args()
    _abortar_se_teste_sem_sandbox(args)

    if args.comando == "resetar-teste":
        from treinamentos_its.ambiente_teste import resetar_sandbox

        resetar_sandbox()
        sys.exit(0)
    if args.comando == "limpar-rodapes":
        from treinamentos_its.limpeza_rodape import limpar_rodapes_arquivo

        sys.exit(0 if limpar_rodapes_arquivo(args.arquivo, dry_run=not args.aplicar) else 1)
    if args.comando == "conferir":
        sys.exit(
            0
            if conferir_matriz(
                args.mes, args.ano, args.pops_vigentes, args.pops_obsoletos
            )
            else 1
        )
    if args.comando == "concluir":
        data_conclusao = None
        if args.data:
            try:
                data_conclusao = datetime.strptime(args.data, "%d/%m/%Y")
            except ValueError:
                parser.error("--data deve usar o formato dd/mm/aaaa.")
        if (args.mes is None) != (args.ano is None):
            parser.error("--mes e --ano devem ser informados juntos (ou nenhum dos dois).")
        if args.mes is not None and not 1 <= args.mes <= 12:
            parser.error("--mes deve estar entre 1 e 12.")
        sys.exit(
            registrar_conclusao(
                args.nome, args.codigo, data_conclusao, args.mes, args.ano
            )
        )
    if args.comando == "sync":
        sys.exit(_run_pipeline_mensal(args.mes, args.ano, args.vigentes))

    parser.print_help()
    log.error("Nenhum comando informado; nada foi executado.")
    sys.exit(2)
