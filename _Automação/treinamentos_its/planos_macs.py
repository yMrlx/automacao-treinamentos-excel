"""Cria e atualiza diretamente o Planos e Macs mensal."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from treinamentos_its.backup import desligar_autosave
from treinamentos_its.config import (
    ABAS_DE_CARGO,
    CABECALHOS_IGNORAR_NORM,
    COL_MATRIZ_CODIGO,
    COL_MATRIZ_VERSAO_ATUAL,
    COL_MATRIZ_VERSAO_NOVA,
    LINHA_INICIO_MATRIZ,
    PASTA_PLANOS_MACS,
)
from treinamentos_its.dominio import (
    MARCACAO_ATRASADO_PRAZO_ANTIGO,
    MARCACAO_ATRASADO_PRAZO_NOVO,
    MARCACAO_ON_TIME,
    MARCACAO_PRESERVAR,
    MARCACAO_STATUS_DESCONHECIDO,
    AvaliacaoMatriz,
    decidir_marcacao_planos,
    esta_atrasado,
    valor_para_quitacao,
    versao_numerica_valida,
    versoes_equivalentes,
)
from treinamentos_its.excel_app import abrir_livro, app_excel
from treinamentos_its.excel_utils import (
    formatar_data_texto,
    normalize_name,
    parse_data,
    sem_acentos,
    versao_para_gravar,
)
from treinamentos_its.logger import log
from treinamentos_its.relatorio import escrever_relatorio_no_rodape, montar_linhas_relatorio


NOME_ABA_MATRIZ = "Matriz - Atualização"


@dataclass(frozen=True)
class MarcacaoAplicada:
    aba: str
    linha: int
    coluna_planejado: int
    planejado: object
    status: str


def _e_cabecalho_treinamentos(celula: object) -> bool:
    if not celula:
        return False
    return sem_acentos(str(celula).strip().casefold()) == "treinamentos"


def dados_tem_conteudo(dados: list | None) -> bool:
    for linha in dados or []:
        if isinstance(linha, list):
            if any(celula not in (None, "") for celula in linha):
                return True
        elif linha not in (None, ""):
            return True
    return False


def _ler_blocos(ws_plan) -> tuple[list[dict], bool]:
    """``(blocos com código, a aba tem algum dado?)`` de uma aba de cargo.

    Cada bloco devolvido é um dict com:

    - ``linha_cod``: linha dos códigos (logo abaixo do cabeçalho "Treinamentos");
    - ``linha_func``: primeira linha de pessoas (cabeçalho + 4);
    - ``linha_fim``: onde o bloco TERMINA, exclusivo — a linha do PRÓXIMO
      cabeçalho "Treinamentos" da aba, tenha ele código ou não. Sem próximo
      cabeçalho, ``ultima_linha + 1``;
    - ``codigos``: coluna (1-based) -> código.

    Só vira bloco o cabeçalho cuja linha de baixo tem código no formato
    ``X-Y-Z`` (regra de sempre). Mas TODO cabeçalho "Treinamentos" encerra o
    bloco de cima: as abas reais têm vários blocos sem código nesse formato
    ("Garantia da Qualidade"/BPF, "Global - Laces", "Ilearn", "Sistemas da
    Qualidade"...). Quando o fim era o próximo bloco COM código, as pessoas
    desses blocos caíam na faixa do bloco de cima e eram lidas nas colunas de
    código dele — o ciclo marcava ``ON TIME`` em célula de outro curso e o
    ``concluir`` gravava ``OK`` em cursos que a pessoa não concluiu.

    O fim do bloco é decidido AQUI e só aqui: ``_atualizar_aba`` e
    ``conclusao._localizar_na_aba`` usam ``linha_fim`` direto, sem recalcular.
    """
    try:
        if ws_plan.api.AutoFilterMode:
            ws_plan.api.AutoFilterMode = False
    except Exception:  # noqa: BLE001
        pass
    ultima_linha = ws_plan.used_range.last_cell.row
    ultima_coluna = ws_plan.used_range.last_cell.column
    dados = ws_plan.range((1, 1), (ultima_linha, ultima_coluna)).value
    if not dados:
        return [], False
    if not isinstance(dados, list):
        dados = [[dados]]
    elif dados and not isinstance(dados[0], list):
        dados = [dados]

    # Linha de TODO cabeçalho "Treinamentos" (com ou sem código), em ordem.
    linhas_cabecalho: list[int] = []
    # (linha do cabeçalho, bloco ainda sem o fim)
    candidatos: list[tuple[int, dict]] = []
    for indice_linha, linha in enumerate(dados):
        if not linha or not any(_e_cabecalho_treinamentos(celula) for celula in linha):
            continue
        linha_treinamentos = indice_linha + 1
        linhas_cabecalho.append(linha_treinamentos)
        linha_codigos = linha_treinamentos + 1
        linha_funcionarios = linha_treinamentos + 4
        if linha_codigos > ultima_linha:
            continue
        valores = dados[linha_codigos - 1]
        if not isinstance(valores, list):
            continue
        codigos: dict[int, str] = {}
        for indice_coluna in range(2, ultima_coluna):
            if indice_coluna >= len(valores):
                break
            valor = valores[indice_coluna]
            if valor:
                codigo = str(valor).strip()
                if len(codigo.split("-")) >= 3:
                    codigos[indice_coluna + 1] = codigo
        if codigos:
            candidatos.append(
                (
                    linha_treinamentos,
                    {
                        "linha_cod": linha_codigos,
                        "linha_func": linha_funcionarios,
                        "codigos": codigos,
                    },
                )
            )

    blocos: list[dict] = []
    for linha_treinamentos, bloco in candidatos:
        proximos = [linha for linha in linhas_cabecalho if linha > linha_treinamentos]
        bloco["linha_fim"] = min(proximos) if proximos else ultima_linha + 1
        blocos.append(bloco)
    return blocos, dados_tem_conteudo(dados)


def mapear_codigos_da_aba(ws_plan) -> list[dict]:
    blocos, tem_conteudo = _ler_blocos(ws_plan)
    if not blocos and tem_conteudo:
        log.warning(
            "Aba '%s': nenhum bloco de 'Treinamentos' reconhecido, apesar de ter dados.",
            ws_plan.name,
        )
    return blocos


def _blocos_e_trava_da_aba(ws_plan) -> tuple[list[dict], str | None]:
    """``(blocos, motivo pra PARAR o ciclo ou None)`` de uma aba de cargo.

    Aba vazia é normal (categoria sem ninguém). Aba COM dados e sem nenhum bloco
    'Treinamentos' reconhecido é trava: a categoria inteira ficaria sem
    marcação e sem rodapé, em silêncio (R2).
    """
    blocos, tem_conteudo = _ler_blocos(ws_plan)
    if not blocos and tem_conteudo:
        return blocos, (
            f"aba '{ws_plan.name}' tem dados mas nenhum bloco 'Treinamentos' "
            f"reconhecido - confira o cabeçalho do bloco (deve ser a palavra "
            f"'Treinamentos')"
        )
    return blocos, None


def travas_das_abas_de_cargo(wb) -> list[str]:
    """Motivos pelos quais as abas de cargo de ``wb`` fazem o ciclo PARAR.

    Checagem ÚNICA usada pelo ``sync`` (``sync.avaliar_arquivo_matriz`` e, de
    novo, ``atualizar_planos_e_macs`` antes da primeira escrita) e pelo
    ``conferir`` (que usa o mesmo ``avaliar_arquivo_matriz``) — assim o
    ``conferir`` não diz OK pra algo que vai parar o ciclo. Não escreve células
    nem salva o workbook, mas ``_ler_blocos`` desliga o AutoFilter da aba em
    memória antes da leitura. Lista vazia = pode seguir.
    """
    nomes = {ws.name for ws in wb.sheets}
    travas: list[str] = []
    faltantes = [aba for aba in ABAS_DE_CARGO if aba not in nomes]
    if faltantes:
        travas.append(
            "aba(s) de cargo não encontrada(s) no arquivo: "
            + ", ".join(faltantes)
            + " - confira o nome da aba no Excel (tem que bater com ABAS_DE_CARGO "
            "do config.py)"
        )
    for nome_aba in ABAS_DE_CARGO:
        if nome_aba not in nomes:
            continue
        _blocos, trava = _blocos_e_trava_da_aba(wb.sheets[nome_aba])
        if trava:
            travas.append(trava)
    return travas


def codigos_dos_blocos(blocos: list[dict]) -> set[str]:
    return {
        str(codigo).strip()
        for bloco in blocos or []
        for codigo in (bloco.get("codigos") or {}).values()
        if str(codigo or "").strip()
    }


def filtrar_verificar_da_aba(
    verificar_codigos: list[str] | None, blocos: list[dict]
) -> list[str]:
    da_aba = codigos_dos_blocos(blocos)
    return list(
        dict.fromkeys(
            str(codigo).strip()
            for codigo in (verificar_codigos or [])
            if str(codigo or "").strip() in da_aba
        )
    )


def codigos_versionados_sem_bloco(
    alteracoes: dict, codigos_encontrados: set[str]
) -> list[str]:
    """Códigos versionados que ninguém tem em nenhum bloco do Planos."""
    return sorted(set(alteracoes or {}) - set(codigos_encontrados or set()))


def avisar_versionados_sem_bloco(
    alteracoes: dict, codigos_encontrados: set[str]
) -> list[str]:
    """Registra os órfãos sem transformar o caso em falha do ciclo."""
    orfaos = codigos_versionados_sem_bloco(alteracoes, codigos_encontrados)
    for codigo in orfaos:
        log.warning(
            "%s versionou mas ninguém tem esse curso no Planos; a versão nova será "
            "registrada na Matriz e o ciclo seguirá normalmente.",
            codigo,
        )
    return orfaos


def avisar_verificar_sem_bloco(
    verificar: list[str] | None, codigos_encontrados: set[str], nome_arquivo: str
) -> list[str]:
    """Códigos ``VERIFICAR`` da Matriz que não aparecem em nenhum bloco.

    Contra falha silenciosa: a linha "POP obsoletado ou renomeado" só vai pro
    rodapé das abas onde o código aparece. Se ele não aparece em nenhuma, o
    único lugar onde alguém fica sabendo é este aviso (além do aviso geral
    "VERIFICAR" da Matriz). Não para o ciclo.
    """
    orfaos = [
        codigo
        for codigo in dict.fromkeys(
            str(codigo).strip() for codigo in (verificar or []) if str(codigo or "").strip()
        )
        if codigo not in (codigos_encontrados or set())
    ]
    if orfaos:
        log.warning(
            "%d código(s) 'VERIFICAR' da Matriz não aparecem em nenhuma aba de cargo "
            "de '%s' - nenhuma linha de rodapé foi escrita pra eles (o aviso fica só "
            "aqui no log; confira a Matriz): %s",
            len(orfaos),
            nome_arquivo,
            ", ".join(orfaos),
        )
    return orfaos


def criar_planos_do_mes(arquivo_origem, arquivo_saida) -> bool:
    """Copia o mês anterior sem nunca usar SaveAs."""
    arquivo_origem = Path(arquivo_origem)
    arquivo_saida = Path(arquivo_saida)
    if arquivo_saida.exists():
        return True
    if not arquivo_origem.exists():
        log.error("Arquivo de origem não encontrado: %s", arquivo_origem.resolve())
        return False
    try:
        arquivo_saida.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(arquivo_origem, arquivo_saida)
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao copiar arquivo de origem: %s", erro)
        return False
    log.info("Criado: %s (cópia de '%s')", arquivo_saida.name, arquivo_origem.name)
    return True


def _pintar_status(celula_planejado, celula_status, status: str, planejado: object) -> None:
    if status == "ON TIME":
        celula_planejado.value = formatar_data_texto(planejado)
        celula_planejado.font.color = (0, 0, 0)
        celula_planejado.color = (105, 217, 105)
        celula_status.value = status
        celula_status.color = (255, 255, 255)
        celula_status.font.color = (0, 128, 0)
        return
    if planejado is not None:
        celula_planejado.value = formatar_data_texto(planejado)
        celula_planejado.font.color = (0, 0, 0)
        celula_planejado.color = None
    celula_status.value = "ATRASADO"
    celula_status.color = (255, 255, 255)
    celula_status.font.color = (255, 0, 0)


def _atualizar_aba(
    ws_plan,
    avaliacao: AvaliacaoMatriz,
    meses: dict,
    hoje: datetime,
) -> tuple[set[str], list[MarcacaoAplicada], list[dict], list[str]]:
    """Atualiza uma aba usando somente o estado que já existe nela.

    As linhas ``<NOME> - <código> ATRASADO`` do rodapé usam ``normalize_name``
    (CAIXA ALTA, sem acento), o formato histórico do entregável — que é
    imutável. É também o que deixa o ``escrever_relatorio_no_rodape``
    reconhecer um bloco idêntico já escrito (comparação exata, linha a linha).
    """
    blocos, trava = _blocos_e_trava_da_aba(ws_plan)
    if trava:
        # Defesa em profundidade: `atualizar_planos_e_macs` já checa todas as
        # abas antes da primeira escrita, com a mesma função.
        raise RuntimeError(trava)
    codigos_aba = codigos_dos_blocos(blocos)
    alteracoes_aba = {
        codigo: dados
        for codigo, dados in avaliacao.alteracoes.items()
        if codigo in codigos_aba
    }
    verificar_aba = filtrar_verificar_da_aba(avaliacao.verificar, blocos)
    marcacoes: list[MarcacaoAplicada] = []
    atrasados: list[dict] = []
    desconhecidos: list[str] = []

    for bloco in blocos:
        # O fim vem pronto do `_ler_blocos` (próximo cabeçalho "Treinamentos",
        # com ou sem código) — não recalcular aqui.
        for linha in range(bloco["linha_func"], bloco["linha_fim"]):
            nome = ws_plan.range(linha, 2).value
            if not nome:
                continue
            nome_norm = normalize_name(nome)
            if nome_norm in CABECALHOS_IGNORAR_NORM:
                continue
            for coluna, codigo in bloco["codigos"].items():
                celula_planejado = ws_plan.range(linha, coluna)
                celula_status = ws_plan.range(linha, coluna + 1)
                status_original = celula_status.value
                planejado_original = celula_planejado.value

                if codigo in avaliacao.alteracoes:
                    prazo = avaliacao.prazos[codigo]
                    decisao = decidir_marcacao_planos(
                        status_original, planejado_original, prazo, hoje
                    )
                    if decisao == MARCACAO_ON_TIME:
                        _pintar_status(celula_planejado, celula_status, "ON TIME", prazo)
                        marcacoes.append(
                            MarcacaoAplicada(ws_plan.name, linha, coluna, prazo, "ON TIME")
                        )
                    elif decisao == MARCACAO_ATRASADO_PRAZO_NOVO:
                        _pintar_status(celula_planejado, celula_status, "ATRASADO", prazo)
                        marcacoes.append(
                            MarcacaoAplicada(ws_plan.name, linha, coluna, prazo, "ATRASADO")
                        )
                        atrasados.append({"nome": nome_norm, "codigo": codigo})
                    elif decisao == MARCACAO_ATRASADO_PRAZO_ANTIGO:
                        # ATRASADO continua com a data antiga; ON TIME vencido
                        # muda apenas o status.
                        _pintar_status(
                            celula_planejado, celula_status, "ATRASADO", None
                        )
                        marcacoes.append(
                            MarcacaoAplicada(
                                ws_plan.name,
                                linha,
                                coluna,
                                planejado_original,
                                "ATRASADO",
                            )
                        )
                        atrasados.append({"nome": nome_norm, "codigo": codigo})
                    elif decisao == MARCACAO_STATUS_DESCONHECIDO:
                        desconhecidos.append(
                            f"aba '{ws_plan.name}' | pessoa '{str(nome).strip()}' | "
                            f"{codigo} | status={status_original!r}"
                        )
                    elif decisao != MARCACAO_PRESERVAR:
                        raise RuntimeError(f"decisão de status desconhecida: {decisao}")
                elif esta_atrasado(status_original, planejado_original, hoje):
                    _pintar_status(
                        celula_planejado, celula_status, "ATRASADO", None
                    )
                    marcacoes.append(
                        MarcacaoAplicada(
                            ws_plan.name,
                            linha,
                            coluna,
                            planejado_original,
                            "ATRASADO",
                        )
                    )
                    atrasados.append({"nome": nome_norm, "codigo": codigo})

    linhas_relatorio = montar_linhas_relatorio(
        meses,
        alteracoes_aba,
        atrasados,
        rotulo_versionado="Versionado treinamento",
        verificar=verificar_aba,
    )
    escrever_relatorio_no_rodape(
        ws_plan, linhas_relatorio, titulo_negrito=True, titulo_tamanho=12
    )
    log.info(
        "Aba '%s': %d marcação(ões), %d versionamento(s), %d VERIFICAR.",
        ws_plan.name,
        len(marcacoes),
        len(alteracoes_aba),
        len(verificar_aba),
    )
    return codigos_aba, marcacoes, atrasados, desconhecidos


def _quitar_matriz_seletivamente(
    wb_planos, linhas_quitacao: dict[str, int]
) -> int:
    nomes = {aba.name for aba in wb_planos.sheets}
    if NOME_ABA_MATRIZ not in nomes:
        raise RuntimeError(f"Aba '{NOME_ABA_MATRIZ}' não encontrada.")
    ws = wb_planos.sheets[NOME_ABA_MATRIZ]
    quitadas = 0
    for codigo, linha in linhas_quitacao.items():
        codigo_lido = str(ws.range(linha, COL_MATRIZ_CODIGO).value or "").strip()
        celula_f = ws.range(linha, COL_MATRIZ_VERSAO_NOVA)
        valor_f = celula_f.value
        valor_e = ws.range(linha, COL_MATRIZ_VERSAO_ATUAL).value
        formula_f = str(getattr(celula_f, "formula", "") or "").strip()
        valor = valor_para_quitacao(valor_f)
        if (
            codigo_lido != codigo
            or valor is None
            or not versao_numerica_valida(valor_e)
            or not formula_f.startswith("=")
        ):
            raise RuntimeError(
                f"A linha {linha} de {codigo} mudou antes de registrar a versão nova na "
                f"Matriz; o ciclo será desfeito."
            )
        ws.range(linha, COL_MATRIZ_VERSAO_ATUAL).value = versao_para_gravar(valor)
        quitadas += 1
    log.info(
        "Versões novas registradas na Matriz (a coluna E recebeu a versão da F): "
        "%d linha(s).",
        quitadas,
    )
    return quitadas


def _planejado_confere(lido: object, esperado: object) -> bool:
    data_lida = parse_data(lido)
    data_esperada = parse_data(esperado)
    if data_lida is not None or data_esperada is not None:
        return (
            data_lida is not None
            and data_esperada is not None
            and data_lida.date() == data_esperada.date()
        )
    return str(lido or "").replace("'", "").strip() == str(esperado or "").replace("'", "").strip()


def _reabrir_e_conferir(
    arquivo_saida: Path,
    avaliacao: AvaliacaoMatriz,
    marcacoes: list[MarcacaoAplicada],
) -> bool:
    with app_excel() as app:
        wb = abrir_livro(app, arquivo_saida, read_only=True)
        try:
            ws_matriz = wb.sheets[NOME_ABA_MATRIZ]
            for codigo, linha in avaliacao.linhas_quitacao.items():
                valor_e = ws_matriz.range(linha, COL_MATRIZ_VERSAO_ATUAL).value
                valor_f = ws_matriz.range(linha, COL_MATRIZ_VERSAO_NOVA).value
                if not versoes_equivalentes(valor_e, valor_f):
                    log.error(
                        "A versão nova de %s não permaneceu salva na Matriz depois de "
                        "reabrir.",
                        codigo,
                    )
                    return False
            for item in marcacoes[:10]:
                ws = wb.sheets[item.aba]
                if str(ws.range(item.linha, item.coluna_planejado + 1).value or "").strip().upper() != item.status:
                    log.error("Amostra de marcação não permaneceu salva em %s.", item.aba)
                    return False
                if not _planejado_confere(
                    ws.range(item.linha, item.coluna_planejado).value, item.planejado
                ):
                    log.error("Prazo da amostra não permaneceu salvo em %s.", item.aba)
                    return False
        finally:
            wb.close()
    log.info(
        "Planos e Macs conferido após reabrir: as versões novas da Matriz e uma "
        "amostra das marcações estão salvas."
    )
    return True


def atualizar_planos_e_macs(
    avaliacao: AvaliacaoMatriz,
    meses: dict,
    arquivo_saida: str | Path | None = None,
    *,
    hoje: datetime | None = None,
) -> bool:
    """Marca os blocos, escreve rodapé, quita por último e reabre."""
    arquivo_saida = Path(arquivo_saida) if arquivo_saida else PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    )
    hoje = hoje or datetime.now()
    if not arquivo_saida.exists():
        log.error("Arquivo não encontrado: %s", arquivo_saida.resolve())
        return False
    if not avaliacao.pode_executar:
        log.error("A Matriz tem bloqueios; o Planos e Macs não será alterado.")
        return False

    codigos_encontrados: set[str] = set()
    marcacoes: list[MarcacaoAplicada] = []
    desconhecidos: list[str] = []
    try:
        with app_excel() as app:
            wb = abrir_livro(app, arquivo_saida)
            try:
                desligar_autosave(wb, descricao=f"'{arquivo_saida.name}' durante o ciclo")
                # Mesma checagem que o `conferir` e o `avaliar_arquivo_matriz`
                # fazem; repetida aqui, no arquivo aberto pra escrita, ANTES da
                # primeira escrita em qualquer aba.
                travas = travas_das_abas_de_cargo(wb)
                if travas:
                    for trava in travas:
                        log.error("O ciclo vai PARAR por isto: %s", trava)
                    return False
                for nome_aba in ABAS_DE_CARGO:
                    codigos, novas_marcacoes, _atrasados, novos_desconhecidos = _atualizar_aba(
                        wb.sheets[nome_aba], avaliacao, meses, hoje
                    )
                    codigos_encontrados.update(codigos)
                    marcacoes.extend(novas_marcacoes)
                    desconhecidos.extend(novos_desconhecidos)

                avisar_versionados_sem_bloco(
                    avaliacao.alteracoes, codigos_encontrados
                )
                avisar_verificar_sem_bloco(
                    avaliacao.verificar, codigos_encontrados, arquivo_saida.name
                )

                # Última mutação do ciclo.
                _quitar_matriz_seletivamente(wb, avaliacao.linhas_quitacao)
                wb.save()
            finally:
                wb.close()
        if not _reabrir_e_conferir(arquivo_saida, avaliacao, marcacoes):
            return False
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao atualizar o Planos e Macs: %s", erro)
        return False

    log.info("Planos e Macs atualizado com sucesso.")
    # De propósito a ÚLTIMA coisa que o ciclo loga (depois da quitação
    # conferida): é o que a operadora precisa resolver à mão, não pode sumir no meio
    # do log.
    _avisar_status_desconhecidos(arquivo_saida, desconhecidos)
    return True


def _avisar_status_desconhecidos(arquivo: Path, desconhecidos: list[str]) -> None:
    """Lista completa dos status que o ciclo NÃO mexeu por não reconhecer."""
    if not desconhecidos:
        return
    log.warning(
        "%d status desconhecido(s) em '%s' NÃO foram mexidos (o ciclo não adivinha "
        "status) - confira cada um à mão:",
        len(desconhecidos),
        arquivo.name,
    )
    for item in desconhecidos:
        log.warning("  - '%s' | %s", arquivo.name, item)
