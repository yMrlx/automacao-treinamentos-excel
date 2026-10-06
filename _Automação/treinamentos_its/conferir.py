"""Conferência somente leitura do mesmo ciclo executado pelo ``sync``."""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from treinamentos_its.config import PASTA_PLANOS_MACS, PASTA_POPS_SISTEMA
from treinamentos_its.excel_app import abrir_livro, app_excel
from treinamentos_its.excel_utils import localizar_coluna, normalizar_versao
from treinamentos_its.logger import log
from treinamentos_its.sync import avaliar_arquivo_matriz, obter_meses


NOME_POPS_OBSOLETOS = "POPs Obsoletos"
CABECALHO_CODIGO = "Document Number"
CABECALHOS_VERSAO_OBSOLETOS = ("Major Version Number", "Version")


def _pastas_dos_exports() -> tuple[Path, Path]:
    """Onde procurar os exports: ``POPs do Sistema/`` e, como compat, a raiz.

    A subpasta é derivada AGORA da ``PASTA_PLANOS_MACS`` deste módulo (a mesma
    relação que o ``config`` usa), e não da constante pronta: assim quem
    redireciona a pasta do Planos (sandbox, testes) redireciona os exports junto
    e nunca cai, por engano, na pasta real.
    """
    return PASTA_PLANOS_MACS / PASTA_POPS_SISTEMA.name, PASTA_PLANOS_MACS


def _arquivo_mais_recente(prefixo: str):
    """Último ``.xlsx`` cujo nome começa com ``prefixo``, mais recente primeiro.

    Procura em ``Planos e Macs/POPs do Sistema/`` e, se não achar nada lá, na
    raiz de ``Planos e Macs/`` (compat). ``glob`` não é recursivo.
    """
    for pasta in _pastas_dos_exports():
        candidatos = sorted(
            pasta.glob(f"{prefixo}*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if candidatos:
            return candidatos[0]
    return None


def _escolher_arquivo(prefixo: str, arquivo=None):
    """O arquivo indicado pelo operador ou, na falta dele, o mais recente da pasta.

    A escolha aparece no log com todas as letras — indicado, achado sozinho (e
    qual), ou nenhum: a escolha implícita por data de modificação já conferiu
    contra o mês errado em silêncio.
    """
    if arquivo:
        caminho = Path(arquivo)
        if not caminho.exists():
            log.error("Arquivo indicado não existe: %s", caminho)
            return None
        log.info("Arquivo INDICADO pelo operador: %s", caminho)
        return caminho
    escolhido = _arquivo_mais_recente(prefixo)
    if escolhido:
        log.info(
            "Nenhum export '%s' indicado; usando o mais recente da pasta (por data "
            "de modificação): %s",
            prefixo,
            escolhido,
        )
    else:
        pasta_sistema, pasta_raiz = _pastas_dos_exports()
        log.info(
            "Nenhum export '%s ...xlsx' indicado nem encontrado em '%s' (nem na raiz "
            "de '%s').",
            prefixo,
            pasta_sistema,
            pasta_raiz,
        )
    return escolhido


def _ler_versoes_por_codigo(
    app, caminho, nomes_versao: tuple[str, ...]
) -> dict[str, str] | None:
    wb = abrir_livro(app, caminho, read_only=True)
    try:
        ws = wb.sheets[0]
        ultima = ws.used_range.last_cell.row
        ultima_coluna = ws.used_range.last_cell.column
        cabecalho = ws.range((1, 1), (1, ultima_coluna)).value
        if not isinstance(cabecalho, list):
            cabecalho = [cabecalho]
        col_codigo = localizar_coluna(cabecalho, (CABECALHO_CODIGO,))
        col_versao = localizar_coluna(cabecalho, nomes_versao)
        if not col_codigo or not col_versao:
            log.error(
                "Cabeçalho não reconhecido em '%s'; a checagem não foi feita.",
                getattr(caminho, "name", caminho),
            )
            return None
        codigos = ws.range((2, col_codigo), (ultima, col_codigo)).value
        versoes = ws.range((2, col_versao), (ultima, col_versao)).value
        if not isinstance(codigos, list):
            codigos, versoes = [codigos], [versoes]
        return {
            str(codigo).strip(): normalizar_versao(versao)
            for codigo, versao in zip(codigos, versoes)
            if codigo
            and str(codigo).strip().upper().startswith(("DOC-",))
        }
    finally:
        wb.close()


def _conferir_contra_obsoletos(codigos_matriz: list[str], arquivo) -> bool:
    """``True`` se a checagem foi FEITA (com ou sem achado); ``False`` se não deu.

    Arquivo que não abre (corrompido, aberto em outro lugar, Excel que não
    sobe) vira mensagem de erro clara e ``False`` — nunca traceback, nunca OK.
    """
    arquivo = Path(arquivo)
    log.info("Comparando os códigos da Matriz com o export de obsoletos: %s", arquivo)
    try:
        with app_excel() as app:
            obsoletos = _ler_versoes_por_codigo(app, arquivo, CABECALHOS_VERSAO_OBSOLETOS)
    except Exception as erro:  # noqa: BLE001
        log.error(
            "NÃO FOI POSSÍVEL LER o export de obsoletos '%s' (%s: %s). A checagem de "
            "obsoletos NÃO foi feita - confira se o arquivo abre no Excel (não está "
            "corrompido nem aberto em outro lugar) e rode a conferência de novo.",
            arquivo.name,
            type(erro).__name__,
            erro,
        )
        return False
    if obsoletos is None:
        return False
    encontrados = sorted(set(codigos_matriz) & set(obsoletos))
    if encontrados:
        log.warning(
            "%d código(s) da Matriz constam em '%s': %s",
            len(encontrados),
            arquivo.name,
            ", ".join(encontrados),
        )
    else:
        log.info("Nenhum código da Matriz consta no export de obsoletos. OK.")
    return True


def _analisar_e_cruzar(
    arquivo_matriz: Path,
    *,
    pops_vigentes=None,
    pops_obsoletos=None,
    vigentes_ja_importado: bool = False,
) -> bool:
    """Avalia a Matriz (mesma avaliação do ``sync``) e cruza com obsoletos.

    ``pops_obsoletos`` chega já resolvido pelo ``conferir_matriz`` (indicado ou
    o mais recente da pasta); ``None`` = não há export pra cruzar.
    """
    avaliacao = avaliar_arquivo_matriz(arquivo_matriz)
    if avaliacao is None:
        return False

    if pops_vigentes and vigentes_ja_importado:
        log.info(
            "Comparar a Matriz com o mesmo export de vigentes recém-importado seria "
            "tautologia; checagem pulada de propósito."
        )
    else:
        log.info(
            "Sem export de vigentes importado: não há export de 'POPs aprovados' "
            "pra comparar com a Matriz."
        )

    if pops_obsoletos:
        obsoletos_ok = _conferir_contra_obsoletos(
            avaliacao.codigos_matriz,
            Path(pops_obsoletos),
        )
    else:
        log.info("Sem export de obsoletos: checagem de obsoletos pulada.")
        obsoletos_ok = True

    if avaliacao.bloqueios:
        log.error(
            "A conferência terminou com erro: o ciclo vai PARAR por isto (%d bloqueio(s)).",
            len(avaliacao.bloqueios),
        )
        return False
    if not obsoletos_ok:
        log.error(
            "Conferência INCOMPLETA: a checagem contra o export de obsoletos não pôde "
            "ser feita (erro acima). Isto NÃO é um 'está tudo certo'."
        )
        return False
    log.info("Conferência concluída. Nada foi alterado.")
    return True


def conferir_matriz(
    mes: int | None = None,
    ano: int | None = None,
    pops_vigentes=None,
    pops_obsoletos=None,
) -> bool:
    """Simula o mês novo com export ou analisa o anterior como está.

    O export de VIGENTES nunca é pego sozinho da pasta (espec de 25/09: sem ele
    vale a Matriz do anterior como está). O de OBSOLETOS continua como sempre
    foi: o indicado ou, sem indicação, o mais recente de ``POPs do Sistema/``.
    """
    try:
        meses = obter_meses(mes, ano)
    except ValueError as erro:
        log.error("Referência inválida: %s", erro)
        return False

    anterior = PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['anterior']} {meses['ano_ant_relativo']}.xlsx"
    )
    novo = PASTA_PLANOS_MACS / (
        f"Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
    )
    if not anterior.exists():
        log.error("Arquivo não encontrado: %s", anterior.resolve())
        return False

    arquivo_obsoletos = _escolher_arquivo(NOME_POPS_OBSOLETOS, pops_obsoletos)
    if pops_obsoletos and arquivo_obsoletos is None:
        log.error(
            "NÃO FOI POSSÍVEL CONFERIR: o export de obsoletos indicado não existe "
            "(erro acima). Escolha o arquivo certo e tente de novo."
        )
        return False

    if not pops_vigentes:
        log.info(
            "Sem --pops-vigentes: conferindo a Matriz do mês anterior como está, "
            "sem criar cópia persistente."
        )
        return _analisar_e_cruzar(
            anterior, pops_obsoletos=arquivo_obsoletos
        )

    arquivo_vigentes = Path(pops_vigentes)
    if not arquivo_vigentes.exists():
        log.error("Arquivo indicado não existe: %s", arquivo_vigentes)
        return False

    from treinamentos_its.planos_macs import criar_planos_do_mes
    from treinamentos_its.vigentes import importar_vigentes

    origem = novo if novo.exists() else anterior
    if novo.exists():
        log.warning(
            "O Planos e Macs do mês novo já existe; a cópia descartável partirá "
            "DELE para representar o reprocessamento."
        )
    try:
        with TemporaryDirectory(
            prefix="treinamentos_its_conferir_", ignore_cleanup_errors=True
        ) as pasta_temporaria:
            temporario = Path(pasta_temporaria) / (
                f"Conferencia - Planos e Macs {meses['atual']} {meses['ano_atual']}.xlsx"
            )
            log.info("Preparando cópia de trabalho DESCARTÁVEL de '%s'.", origem.name)
            if not criar_planos_do_mes(origem, temporario):
                return False
            if not importar_vigentes(
                arquivo_vigentes,
                temporario,
                destino_copia_restauracao=Path(pasta_temporaria)
                / "_pre-colagem"
                / temporario.name,
            ):
                return False
            resultado = _analisar_e_cruzar(
                temporario,
                pops_vigentes=arquivo_vigentes,
                pops_obsoletos=arquivo_obsoletos,
                vigentes_ja_importado=True,
            )
        log.info("Cópia de trabalho descartada. Nenhum arquivo persistente foi alterado.")
        return resultado
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao preparar a conferência descartável: %s", erro)
        return False
