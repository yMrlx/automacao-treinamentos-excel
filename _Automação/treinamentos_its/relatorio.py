"""Montagem e escrita do relatório do ciclo.

O relatório continua sendo escrito como texto solto, nas primeiras linhas livres
após a tabela de dados de cada aba. A montagem do texto (``montar_linhas_relatorio``,
pura) é separada da escrita no Excel (``escrever_relatorio_no_rodape``) para que a
primeira possa ser testada sem abrir nenhum workbook.
"""

from treinamentos_its.excel_utils import versao_e_numerica
from treinamentos_its.logger import log


def _rotulo_versao(valor: str) -> str:
    """``'2.0'`` -> ``'v2.0'``; ``'VERIFICAR'`` (ou outro texto) -> sem o ``v``."""
    return f"v{valor}" if versao_e_numerica(valor) else str(valor)


def montar_linhas_relatorio(
    meses: dict,
    alteracoes: dict[str, dict[str, str]],
    atrasados: list[dict[str, str]],
    rotulo_versionado: str = "Versionado do treinamento",
    verificar: list[str] | None = None,
) -> list[str]:
    """Monta o texto do relatório de um ciclo, na ordem em que é escrito na aba.

    ``verificar`` (decisão A3, 2026-09-02): códigos com ``VERIFICAR`` na coluna F
    da Matriz (POP obsoletado ou renomeado) — não versionam ninguém, só geram uma
    linha extra pedindo pra um humano conferir. Repete todo mês em que a linha
    continuar ``VERIFICAR`` na Matriz (não é "só uma vez").
    """
    # O rodapé leva o mês do CICLO (= do arquivo criado): o Planos e Macs Junho
    # 2026 feito à mão termina em "Referente Junho 2026" (regra do negócio, 06/10).
    linhas = [f"Referente {meses['atual']} {meses['ano_atual']}"]

    if alteracoes:
        for codigo in sorted(alteracoes):
            dados = alteracoes[codigo]
            linhas.append(
                f"{rotulo_versionado} {codigo} de {_rotulo_versao(dados['v_antiga'])} "
                f"para {_rotulo_versao(dados['v_nova'])}"
            )
    else:
        linhas.append("Não houveram alterações nos POPs")

    for codigo in sorted(verificar or []):
        linhas.append(f"{codigo} — POP obsoletado ou renomeado, para um humano verificar")

    for item in atrasados:
        linhas.append(f"{item['nome']} - {item['codigo']} ATRASADO")

    return linhas


def escrever_relatorio_no_rodape(
    ws,
    linhas: list[str],
    titulo_negrito: bool = False,
    titulo_tamanho: int | None = None,
) -> None:
    """Escreve as linhas do relatório a partir da primeira linha livre da aba.

    Idempotente: se este relatório **inteiro** (todas as linhas, na mesma ordem)
    já aparece no rodapé, nada é reescrito — evita empilhar uma segunda cópia
    quando o mesmo mês é reprocessado. Só a linha-título coincidir não basta: a
    planilha do mês é criada copiando a do mês anterior e pode carregar blocos
    antigos com o mesmo título mas conteúdo diferente, que devem ser preservados.
    """
    if not linhas:
        return

    ultima_linha = ws.used_range.last_cell.row
    valores_col_a = ws.range((1, 1), (ultima_linha, 1)).value
    if not isinstance(valores_col_a, list):
        valores_col_a = [valores_col_a]

    n = len(linhas)
    for inicio in range(len(valores_col_a) - n + 1):
        if valores_col_a[inicio:inicio + n] == linhas:
            log.info(
                "Relatório '%s' idêntico já consta no rodapé da aba '%s'; não foi reescrito.",
                linhas[0],
                ws.name,
            )
            return

    linha_atual = ultima_linha + 2
    while ws.range(linha_atual, 1).value is not None:
        linha_atual += 1

    for indice, texto in enumerate(linhas):
        celula = ws.range(linha_atual, 1)
        celula.value = texto
        if indice == 0:
            if titulo_negrito:
                celula.font.bold = True
            if titulo_tamanho:
                celula.font.size = titulo_tamanho
        linha_atual += 1
