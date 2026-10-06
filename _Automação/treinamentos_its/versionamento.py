"""Prazo de cada versionamento — aba 'Versionamento Mês' (R12, 2026-09-18).

Regra de negócio: **o prazo de um treinamento é a data de aprovação do POP mais
30 dias**, e a data de aprovação vem da aba ``Versionamento Mês`` da planilha
``Planos e Macs`` (a mesma que traz a ``Matriz - Atualização``). Portanto o prazo
é **por treinamento**, não uma data única do ciclo — que era o que o código fazia
antes (``meses["data_atual"]``: a data de hoje no caminho padrão, o dia 1º do mês
no caminho ``--mes/--ano``; nenhum dos dois é a regra de negócio).

Sobre o layout da aba: ele **muda de um mês pro outro** (conferido nos arquivos
de Janeiro a Setembro/2026 no sandbox):

- Janeiro/Fevereiro: cabeçalhos em português (``Número do documento``,
  ``Data de aprovação``);
- Abril, Julho, Agosto, Setembro: em inglês, ``Approved Date`` na coluna **G**;
- Junho: tem **duas** colunas ``Version``, e aí a ``Approved Date`` cai na coluna
  **H** — é por isso que a coluna NUNCA é lida por índice fixo, sempre pelo nome
  do cabeçalho (mesma lição do BUG-1 do ``conferir``);
- Maio: a maioria das datas está como **texto em formato americano**
  (``4/30/2026``) em vez de data de verdade.

Quem não for encontrado sai numa lista explícita pra quem chamou avisar alto —
nada de chutar uma data em silêncio.
"""

from datetime import datetime, timedelta

from treinamentos_its.excel_utils import localizar_coluna, normalizar_cabecalho, parse_data
from treinamentos_its.logger import log

# Nome da aba na "Planos e Macs" (com acento e espaço). A busca tolera
# caixa/acento/espaço, igual ao cabeçalho do bloco "Treinamentos" (B4).
NOME_ABA_VERSIONAMENTO = "Versionamento Mês"

# Cabeçalhos aceitos, em ordem de preferência. Os dois idiomas convivem: os
# arquivos de Janeiro/Fevereiro/2026 estão em português, o resto em inglês.
CABECALHOS_CODIGO = ("Document Number", "Número do documento")
CABECALHOS_DATA_APROVACAO = ("Approved Date", "Data de aprovação")

# "30 dias depois da data de aprovação" (regra de negócio, 2026-09-18).
DIAS_DE_PRAZO = 30

LINHA_INICIO_VERSIONAMENTO = 2  # linha 1 é o cabeçalho


class ErroDeLeituraDoVersionamento(Exception):
    """Não deu pra LER a aba de versionamento (aba ausente, cabeçalho mudou...).

    Existe pra separar dois casos que antes se confundiam num dicionário vazio:

    - "li a aba e ninguém versionou / este código não está lá" -> a vida segue,
      com aviso;
    - "não consegui ler o que preciso" -> o ciclo **não pode** continuar: sem as
      datas de aprovação não existe prazo, e salvar assim mesmo seria o sucesso
      falso que o R2/R7 atacaram (o `.bat` diz "terminou sem erros" e metade dos
      treinamentos ficou sem atualização).
    """


def prazo_do_versionamento(aprovado_em: object, dias: int = DIAS_DE_PRAZO) -> datetime | None:
    """Data de aprovação + ``dias`` — o prazo daquele treinamento. PURA.

    Aceita ``datetime`` ou texto (passa por ``parse_data``, que entende
    ``dd/mm/aaaa``, ISO e ``mm/dd/aaaa`` quando o dia > 12). O horário é zerado:
    prazo é um DIA, e é assim que ele é comparado com ``hoje`` e gravado na
    planilha (``dd/mm/yyyy``).

    Sem data legível -> ``None``. Quem chama decide o que fazer com isso; aqui
    não se inventa data.
    """
    data = parse_data(aprovado_em)
    if data is None:
        return None
    return datetime(data.year, data.month, data.day) + timedelta(days=dias)


# --- datas em TEXTO: qual é o dia e qual é o mês? ----------------------------
#
# A aba de Maio/2026 traz a maioria das datas como texto em formato americano
# ("4/30/2026"). `parse_data` resolve os casos inequívocos (30 não é mês), mas
# "5/1/2026" pode ser 5 de janeiro OU 1º de maio - e agora que o PRAZO sai daqui,
# ler errado desloca uma data de compliance em meses. Então: a convenção da
# coluna é inferida pelos casos inequívocos e só depois aplicada nos ambíguos.
# Sem evidência (ou com evidência contraditória), o valor ambíguo é RECUSADO -
# melhor um aviso alto do que um prazo errado com cara de certo.
FORMATO_DIA_PRIMEIRO = "%d/%m/%Y"
FORMATO_MES_PRIMEIRO = "%m/%d/%Y"


def _numeros_de_data_texto(valor: object) -> tuple[int, int, int] | None:
    """``'4/30/2026'`` -> ``(4, 30, 2026)``. ``None`` se não for data com barras."""
    if not isinstance(valor, str):
        return None
    partes = valor.replace("'", "").strip().split("/")
    if len(partes) != 3:
        return None
    try:
        primeiro, segundo, terceiro = (int(parte) for parte in partes)
    except ValueError:
        return None
    return primeiro, segundo, terceiro


def texto_de_data_e_ambiguo(valor: object) -> bool:
    """True se o texto serve tanto como dd/mm quanto como mm/dd. PURA."""
    numeros = _numeros_de_data_texto(valor)
    if numeros is None:
        return False
    primeiro, segundo, _ano = numeros
    return 1 <= primeiro <= 12 and 1 <= segundo <= 12


def detectar_formato_das_datas(valores: list[object]) -> str | None:
    """Convenção das datas em texto da coluna, deduzida dos casos inequívocos. PURA.

    ``'4/30/2026'`` (30 não é mês) prova mm/dd; ``'30/04/2026'`` prova dd/mm.
    Devolve ``None`` quando não há nenhum caso inequívoco **ou** quando há dos
    dois tipos (planilha misturada) — aí ninguém adivinha nada.
    """
    mes_primeiro = dia_primeiro = False
    for valor in valores or []:
        numeros = _numeros_de_data_texto(valor)
        if numeros is None:
            continue
        primeiro, segundo, _ano = numeros
        if primeiro > 12 and segundo <= 12:
            dia_primeiro = True
        elif segundo > 12 and primeiro <= 12:
            mes_primeiro = True
    if dia_primeiro and mes_primeiro:
        return None
    if dia_primeiro:
        return FORMATO_DIA_PRIMEIRO
    if mes_primeiro:
        return FORMATO_MES_PRIMEIRO
    return None


def montar_datas_aprovacao(
    linhas: list[tuple[object, object]],
) -> tuple[dict[str, datetime], list[str], list[str]]:
    """``[(codigo, data_bruta)]`` -> ``(datas, ilegiveis, ambiguos)``. PURA.

    - Linha sem código é ignorada (a aba vem com centenas de linhas vazias no
      fim: o ``used_range`` do Excel infla).
    - Código repetido: vence a data **mais recente** — na aba ela representa a
      aprovação da versão vigente, e é essa que dispara o treinamento do ciclo.
      (Nos arquivos do sandbox isso acontece em pouquíssimos códigos.)
    - ``ilegiveis``: código com a célula de data vazia ou impossível de ler.
    - ``ambiguos``: a data está em texto que serve como dd/mm **e** como mm/dd e
      a coluna não deu pista de qual é (ver ``detectar_formato_das_datas``).

    Quem cai em ``ilegiveis``/``ambiguos`` fica **fora** do dicionário: quem
    chama grita, em vez de seguir com um prazo inventado.
    """
    linhas = list(linhas or [])
    formato = detectar_formato_das_datas([data_bruta for _codigo, data_bruta in linhas])

    datas: dict[str, datetime] = {}
    ilegiveis: list[str] = []
    ambiguos: list[str] = []
    for codigo, data_bruta in linhas:
        codigo_str = str(codigo or "").strip()
        if not codigo_str:
            continue

        if texto_de_data_e_ambiguo(data_bruta):
            if formato is None:
                if codigo_str not in ambiguos:
                    ambiguos.append(codigo_str)
                continue
            try:
                data = datetime.strptime(str(data_bruta).replace("'", "").strip(), formato)
            except ValueError:
                data = None
        else:
            data = parse_data(data_bruta)

        if data is None:
            if codigo_str not in ilegiveis:
                ilegiveis.append(codigo_str)
            continue
        somente_dia = datetime(data.year, data.month, data.day)
        anterior = datas.get(codigo_str)
        if anterior is None or somente_dia > anterior:
            datas[codigo_str] = somente_dia

    # Código que apareceu duas vezes (uma com data boa, outra não) não é
    # problema: o que vale é ter pelo menos uma data legível.
    ilegiveis = [codigo for codigo in ilegiveis if codigo not in datas]
    ambiguos = [codigo for codigo in ambiguos if codigo not in datas]
    return datas, ilegiveis, ambiguos


def prazos_por_codigo(
    codigos: list[str],
    datas_aprovacao: dict[str, datetime],
    dias: int = DIAS_DE_PRAZO,
) -> tuple[dict[str, datetime], list[str]]:
    """``({codigo: prazo}, codigos_sem_prazo)`` pros códigos pedidos. PURA.

    ``codigos_sem_prazo`` = os que não têm data de aprovação na aba (ou têm data
    ilegível). Eles ficam **fora** do dicionário de propósito: o chamador precisa
    tratar o caso explicitamente, não receber um prazo chutado.
    """
    prazos: dict[str, datetime] = {}
    sem_prazo: list[str] = []
    for codigo in codigos or []:
        codigo_str = str(codigo or "").strip()
        if not codigo_str:
            continue
        prazo = prazo_do_versionamento(datas_aprovacao.get(codigo_str), dias)
        if prazo is None:
            if codigo_str not in sem_prazo:
                sem_prazo.append(codigo_str)
            continue
        prazos[codigo_str] = prazo
    return prazos, sem_prazo


def localizar_aba_versionamento(wb):
    """A aba 'Versionamento Mês' do workbook aberto, tolerando caixa/acento/espaço.

    ``None`` quando não existe (o chamador loga o erro com o nome do arquivo, que
    daqui a gente não tem de forma confiável).
    """
    alvo = normalizar_cabecalho(NOME_ABA_VERSIONAMENTO)
    for aba in wb.sheets:
        if normalizar_cabecalho(aba.name) == alvo:
            return aba
    return None


def ler_datas_aprovacao(wb, nome_arquivo: str = "") -> dict[str, datetime]:
    """``{código: data de aprovação}`` lido da aba 'Versionamento Mês' do ``wb``.

    Toca no Excel (por isso fica separada das funções puras acima), mas usa um
    workbook **já aberto**: o ``sync`` lê a Matriz e a aba de versionamento do
    mesmo arquivo, sem abrir nada a mais.

    **Levanta ``ErroDeLeituraDoVersionamento``** quando a aba não existe ou os
    cabeçalhos não batem. Antes isso devolvia ``{}``, o ``sync`` descartava todos
    os versionamentos "por falta de prazo", salvava e dizia que tinha dado certo:
    sucesso falso com o ciclo pela metade. Aba vazia (só o cabeçalho) é outra
    coisa — foi lida, só não tem dado: devolve ``{}`` com aviso.
    """
    rotulo = nome_arquivo or getattr(wb, "name", "") or "(arquivo)"
    ws = localizar_aba_versionamento(wb)
    if ws is None:
        raise ErroDeLeituraDoVersionamento(
            f"Aba '{NOME_ABA_VERSIONAMENTO}' não encontrada em '{rotulo}' - sem ela não "
            f"dá pra calcular o prazo (data de aprovação + {DIAS_DE_PRAZO} dias). Abas "
            f"disponíveis: {', '.join(aba.name for aba in wb.sheets)}"
        )

    ultima_linha = ws.used_range.last_cell.row
    ultima_coluna = ws.used_range.last_cell.column
    cabecalho = ws.range((1, 1), (1, ultima_coluna)).value
    if not isinstance(cabecalho, list):
        cabecalho = [cabecalho]

    col_codigo = localizar_coluna(cabecalho, CABECALHOS_CODIGO)
    col_data = localizar_coluna(cabecalho, CABECALHOS_DATA_APROVACAO)
    if not col_codigo or not col_data:
        raise ErroDeLeituraDoVersionamento(
            f"Não achei a(s) coluna(s) esperada(s) no cabeçalho da aba '{ws.name}' de "
            f"'{rotulo}' (código: "
            f"{' ou '.join(repr(nome) for nome in CABECALHOS_CODIGO)} -> "
            f"{col_codigo or 'não encontrada'}; data de aprovação: "
            f"{' ou '.join(repr(nome) for nome in CABECALHOS_DATA_APROVACAO)} -> "
            f"{col_data or 'não encontrada'}). O layout da aba mudou e o prazo NÃO pôde "
            f"ser calculado. Cabeçalho lido: "
            f"{', '.join(str(titulo) for titulo in cabecalho if titulo)}"
        )

    if ultima_linha < LINHA_INICIO_VERSIONAMENTO:
        log.warning("Aba '%s' de '%s' está vazia (só o cabeçalho).", ws.name, rotulo)
        return {}

    codigos = ws.range(
        (LINHA_INICIO_VERSIONAMENTO, col_codigo), (ultima_linha, col_codigo)
    ).value
    datas_brutas = ws.range(
        (LINHA_INICIO_VERSIONAMENTO, col_data), (ultima_linha, col_data)
    ).value
    if not isinstance(codigos, list):
        codigos, datas_brutas = [codigos], [datas_brutas]

    datas, ilegiveis, ambiguos = montar_datas_aprovacao(list(zip(codigos, datas_brutas)))
    log.info(
        "Aba '%s' de '%s': código na coluna %d ('%s'), data de aprovação na coluna %d "
        "('%s'); %d código(s) com data legível.",
        ws.name,
        rotulo,
        col_codigo,
        cabecalho[col_codigo - 1],
        col_data,
        cabecalho[col_data - 1],
        len(datas),
    )
    if not datas:
        log.warning(
            "A aba '%s' de '%s' foi lida, mas NENHUM código tem data de aprovação. "
            "Qualquer versionamento deste ciclo vai ficar sem prazo.",
            ws.name,
            rotulo,
        )
    if ilegiveis:
        # Não é fatal (pode ser POP que não versiona neste ciclo), mas é
        # exatamente o tipo de coisa que vira falha silenciosa se ninguém vir.
        log.warning(
            "%d código(s) da aba '%s' estão sem data de aprovação legível na coluna "
            "'%s' - se algum deles versionar, vai ficar SEM prazo: %s",
            len(ilegiveis),
            ws.name,
            cabecalho[col_data - 1],
            ", ".join(ilegiveis[:20]) + (" ..." if len(ilegiveis) > 20 else ""),
        )
    if ambiguos:
        log.warning(
            "%d código(s) da aba '%s' têm a data de aprovação em TEXTO AMBÍGUO (ex.: "
            "'5/1/2026' - 5 de janeiro ou 1º de maio?) e a coluna não deixa claro qual "
            "é a convenção. A data foi RECUSADA de propósito (chutar erraria o prazo "
            "em meses): %s. Formate a coluna '%s' como data de verdade no Excel.",
            len(ambiguos),
            ws.name,
            ", ".join(ambiguos[:20]) + (" ..." if len(ambiguos) > 20 else ""),
            cabecalho[col_data - 1],
        )
    return datas
