"""Faxina de rodapés duplicados (T5).

Antes de existir a guarda de reprocessamento (T2), rodar o mesmo mês mais de
uma vez empilhava um bloco de rodapé novo por cima do(s) anterior(es) — mesma
aba, mesmo título ``"Referente <mês> <ano>"``, conteúdo diferente (a lista de
ATRASADO muda entre rodadas; a idempotência de ``relatorio.escrever_relatorio_no_rodape``
só bloqueia bloco 100% idêntico). É o caso real relatado pelo responsável na aba
``Gerente`` do ``Planos e Macs Julho 2026.xlsx``.

Regra combinada com o negócio (2026-09-03): o rodapé é a parte mais importante
do arquivo — tem que manter o **log completo de todos os cursos versionados**
daquela aba. A limpeza não é "apaga o mais velho", é **mesclar**:

- Linhas "Versionado ...": união de todas as ocorrências dos blocos duplicados
  do mesmo título, sem repetir código.
- Linhas "... ATRASADO": usa a lista do bloco MAIS RECENTE (mais próximo do fim
  do arquivo) — é a mais atual.
- Linhas "<código> — POP obsoletado ou renomeado, ...": mesma regra das
  "Versionado" (união, sem repetir código) — é o mesmo tipo de "log que persiste
  enquanto for verdade".
- Sem nenhum "Versionado" sobrando: mantém "Não houveram alterações nos POPs".
- Linhas que não batem com nenhum padrão conhecido: mantidas (união por texto
  exato, sem duplicar) — pra não perder informação por via das dúvidas.
- Resultado: UM bloco por título, na posição do último bloco (os anteriores
  somem, sem deixar buraco). O espaçamento entre blocos é normalizado pra
  exatamente 1 linha em branco, igual ao jeito que `escrever_relatorio_no_rodape`
  já escreve.

``mesclar_blocos_duplicados`` é pura (lista de strings/None -> lista de
strings/None) e testável sem Excel. ``aplicar_mesclagem_na_aba``/
``limpar_rodapes_arquivo`` fazem o trabalho sujo (leem/escrevem células).
"""

from __future__ import annotations

from pathlib import Path

from treinamentos_its.logger import log

TITULO_PREFIXO = "Referente "
NAO_HOUVE_ALTERACAO = "Não houveram alterações nos POPs"
SUFIXO_VERIFICAR = " — POP obsoletado ou renomeado, para um humano verificar"
SUFIXO_ATRASADO = " ATRASADO"


def _e_titulo(valor: object) -> bool:
    return isinstance(valor, str) and valor.startswith(TITULO_PREFIXO)


def _extrair_codigo_versionado(linha: str) -> str | None:
    """Código de uma linha "<rótulo> <código> de <v1> para <v2>", ou ``None``.

    Não fixa o texto do rótulo (varia: "Versionado do treinamento" no
    rótulos históricos e atuais de versionamento) — só exige que a
    linha comece com "Versionado" e tenha o miolo " de ... para ...".
    """
    if not linha.startswith("Versionado") or " de " not in linha or " para " not in linha:
        return None
    antes_de = linha.split(" de ", 1)[0]
    partes = antes_de.rsplit(" ", 1)
    if len(partes) != 2 or not partes[1]:
        return None
    return partes[1]


def _extrair_codigo_verificar(linha: str) -> str | None:
    if linha.endswith(SUFIXO_VERIFICAR):
        codigo = linha[: -len(SUFIXO_VERIFICAR)]
        return codigo or None
    return None


def _linhas_uteis(linhas: list) -> list[str]:
    """Filtra só o texto real de um bloco (tira o ``None`` da linha em branco
    que separa este bloco do próximo)."""
    return [linha for linha in linhas if isinstance(linha, str)]


def _mesclar_grupo(titulo: str, grupos: list[list]) -> list[str]:
    """Mescla N ocorrências (na ordem em que aparecem no arquivo) do mesmo título."""
    versionados: dict[str, str] = {}
    verificar: dict[str, str] = {}
    outras_vistas: list[str] = []
    outras_set: set[str] = set()

    for grupo in grupos:
        for linha in grupo[1:]:  # [0] é o título
            if not isinstance(linha, str):
                continue
            if linha == NAO_HOUVE_ALTERACAO:
                continue  # decidido no final, conforme sobrar ou não "Versionado"
            if linha.endswith(SUFIXO_ATRASADO):
                continue  # ATRASADO só entra do bloco mais recente, tratado abaixo

            codigo_v = _extrair_codigo_versionado(linha)
            if codigo_v is not None:
                versionados.setdefault(codigo_v, linha)
                continue

            codigo_verif = _extrair_codigo_verificar(linha)
            if codigo_verif is not None:
                verificar.setdefault(codigo_verif, linha)
                continue

            if linha not in outras_set:
                outras_set.add(linha)
                outras_vistas.append(linha)

    atrasados_do_ultimo = [
        linha for linha in grupos[-1][1:] if isinstance(linha, str) and linha.endswith(SUFIXO_ATRASADO)
    ]

    linhas_finais = [titulo]
    if versionados:
        for codigo in sorted(versionados):
            linhas_finais.append(versionados[codigo])
    else:
        linhas_finais.append(NAO_HOUVE_ALTERACAO)
    for codigo in sorted(verificar):
        linhas_finais.append(verificar[codigo])
    linhas_finais.extend(outras_vistas)
    linhas_finais.extend(atrasados_do_ultimo)
    return linhas_finais


def mesclar_blocos_duplicados(linhas_aba: list) -> list:
    """Mescla blocos de rodapé com o mesmo título ``"Referente <mês> <ano>"``.

    ``linhas_aba`` é a coluna A inteira de uma aba, célula a célula (o que
    ``ws.range((1,1),(ultima_linha,1)).value`` devolve). Devolve uma lista nova
    (não muda ``linhas_aba``); aba sem título nenhum, ou só com títulos únicos,
    volta EXATAMENTE igual (mesma identidade de conteúdo).
    """
    indices_titulo = [i for i, v in enumerate(linhas_aba) if _e_titulo(v)]
    if not indices_titulo:
        return list(linhas_aba)

    blocos = []
    for pos, inicio in enumerate(indices_titulo):
        fim = indices_titulo[pos + 1] if pos + 1 < len(indices_titulo) else len(linhas_aba)
        blocos.append({"titulo": linhas_aba[inicio], "linhas": linhas_aba[inicio:fim]})

    por_titulo: dict[str, list[int]] = {}
    for i, bloco in enumerate(blocos):
        por_titulo.setdefault(bloco["titulo"], []).append(i)

    if all(len(indices) == 1 for indices in por_titulo.values()):
        return list(linhas_aba)  # nada duplicado - não mexe em nada

    indices_a_remover: set[int] = set()
    conteudo_final: dict[int, list[str]] = {}
    for titulo, indices in por_titulo.items():
        if len(indices) == 1:
            conteudo_final[indices[0]] = _linhas_uteis(blocos[indices[0]]["linhas"])
            continue
        grupos = [blocos[i]["linhas"] for i in indices]
        conteudo_final[indices[-1]] = _mesclar_grupo(titulo, grupos)
        indices_a_remover.update(indices[:-1])

    resultado = list(linhas_aba[: indices_titulo[0]])
    primeiro_bloco_escrito = True
    for i, bloco in enumerate(blocos):
        if i in indices_a_remover:
            continue
        if not primeiro_bloco_escrito:
            resultado.append(None)  # 1 linha em branco entre blocos, igual ao original
        primeiro_bloco_escrito = False
        resultado.extend(conteudo_final[i])
    return resultado


def _formatos_por_titulo(ws, valores: list, indices_titulo: list[int]) -> dict[str, tuple[bool, float | None]]:
    """Formatação (negrito, tamanho) da célula de cada título, capturada ANTES
    de mexer em qualquer coisa. Em ordem crescente, então título repetido fica
    com o formato da ÚLTIMA ocorrência (a que efetivamente sobrevive)."""
    formatos: dict[str, tuple[bool, float | None]] = {}
    for indice in indices_titulo:
        titulo = valores[indice]
        celula = ws.range(indice + 1, 1)
        try:
            formatos[titulo] = (bool(celula.font.bold), celula.font.size)
        except Exception as erro:  # noqa: BLE001
            log.warning("Não consegui ler a formatação do título '%s': %s", titulo, erro)
            formatos[titulo] = (False, None)
    return formatos


def aplicar_mesclagem_na_aba(ws) -> bool:
    """Aplica ``mesclar_blocos_duplicados`` numa aba (xlwings) já aberta p/ escrita.

    Só mexe na coluna A a partir do primeiro título "Referente ..." - a área de
    dados (nomes, códigos, status...) antes dele não é tocada. Devolve ``True``
    se reescreveu algo, ``False`` se a aba já estava OK (nenhum bloco duplicado).
    """
    ultima_linha = ws.used_range.last_cell.row
    if ultima_linha < 1:
        return False
    valores = ws.range((1, 1), (ultima_linha, 1)).value
    if not isinstance(valores, list):
        valores = [valores]

    indices_titulo = [i for i, v in enumerate(valores) if _e_titulo(v)]
    if not indices_titulo:
        return False

    novo = mesclar_blocos_duplicados(valores)
    if novo == valores:
        return False

    formatos = _formatos_por_titulo(ws, valores, indices_titulo)

    primeira_linha_rodape = indices_titulo[0] + 1  # 1-based
    conteudo_novo_rodape = novo[indices_titulo[0]:]
    for offset, texto in enumerate(conteudo_novo_rodape):
        linha_excel = primeira_linha_rodape + offset
        celula = ws.range(linha_excel, 1)
        celula.value = texto
        if _e_titulo(texto):
            negrito, tamanho = formatos.get(texto, (False, None))
            celula.font.bold = negrito
            if tamanho:
                celula.font.size = tamanho

    linha_fim_novo = primeira_linha_rodape + len(conteudo_novo_rodape) - 1
    for linha_excel in range(linha_fim_novo + 1, ultima_linha + 1):
        ws.range(linha_excel, 1).value = None

    return True


def _resumo_diferenca(nome_aba: str, antes: list, depois: list) -> str:
    """Texto curto (pro log/relatório) do que mudou numa aba."""
    antes_titulos = [v for v in antes if _e_titulo(v)]
    depois_titulos = [v for v in depois if _e_titulo(v)]
    return (
        f"Aba '{nome_aba}': {len(antes_titulos)} bloco(s) de rodapé -> "
        f"{len(depois_titulos)} bloco(s) ({len(antes)} -> {len(depois)} linha(s) na coluna A)."
    )


def limpar_rodapes_arquivo(caminho: str | Path, *, dry_run: bool = True) -> bool:
    """Abre um ``.xlsx`` e mescla os blocos de rodapé duplicados de cada aba.

    ``dry_run=True`` (padrão): abre **read-only**, só loga o que mudaria em
    cada aba - NÃO grava nada. ``dry_run=False``: aplica de verdade e salva
    (in-place - mesmo padrão anti-`SaveAs` do resto do projeto, ver B-QA-1).

    Retorna ``False`` só em erro de verdade (arquivo não encontrado, exceção ao
    processar). "Não tinha nada pra mesclar" é sucesso (``True``), só sem
    alterar nada.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        log.error("Arquivo não encontrado: %s", caminho.resolve())
        return False

    from treinamentos_its.excel_app import abrir_livro, app_excel

    try:
        with app_excel() as app:
            wb = abrir_livro(app, caminho, read_only=dry_run)
            try:
                alguma_aba_mudou = False
                for ws in wb.sheets:
                    ultima_linha = ws.used_range.last_cell.row
                    if ultima_linha < 1:
                        continue
                    valores = ws.range((1, 1), (ultima_linha, 1)).value
                    if not isinstance(valores, list):
                        valores = [valores]
                    novo = mesclar_blocos_duplicados(valores)
                    if novo == valores:
                        continue

                    alguma_aba_mudou = True
                    log.info("%s", _resumo_diferenca(ws.name, valores, novo))
                    if dry_run:
                        log.info(
                            "Aba '%s': DRY-RUN, nada foi gravado. Rode com --aplicar pra "
                            "gravar de verdade.",
                            ws.name,
                        )
                    else:
                        aplicar_mesclagem_na_aba(ws)

                if not alguma_aba_mudou:
                    log.info("Nenhum bloco de rodapé duplicado em %s.", caminho.name)
                elif not dry_run:
                    wb.save()
                    log.info("Arquivo salvo: %s", caminho.resolve())
            finally:
                wb.close()
        return True
    except Exception as erro:  # noqa: BLE001
        log.exception("Erro ao limpar rodapés de %s: %s", caminho.name, erro)
        return False
