"""Regras de negócio da automação, sem dependência de Excel.

Estas funções são puras (mesma entrada -> mesma saída, sem abrir workbook) para
poderem ser testadas isoladamente. Reúnem a lógica que antes estava duplicada
entre os módulos que fazem a leitura/escrita no Excel.
"""

import math
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta

from treinamentos_its.config import MARCADOR_VERIFICAR
from treinamentos_its.excel_utils import (
    normalize_name,
    normalizar_versao,
    parse_data,
    versao_e_numerica,
)


def versao_numerica_valida(valor: object) -> bool:
    """True apenas para número de versão finito e não negativo.

    Além de texto como ``VERIFICAR``/``N/A``, rejeita ``NaN``, infinito e números
    negativos. Estes últimos podem aparecer quando o Excel expõe um código COM de
    erro como inteiro; tratá-los como versão quitável destruiria a última versão
    válida da coluna E.
    """
    if isinstance(valor, bool):
        return False
    normalizada = normalizar_versao(valor)
    if not versao_e_numerica(normalizada):
        return False
    try:
        numero = float(normalizada)
    except (TypeError, ValueError):
        return False
    return math.isfinite(numero) and numero >= 0


def valor_para_quitacao(versao_nova: object) -> object | None:
    """Valor de F que pode ser copiado para E; ``None`` significa não tocar E.

    O valor bruto é devolvido de propósito: a quitação equivale a copiar e colar
    **como valor**, preservando se o Excel guarda a versão como número ou texto.
    """
    return versao_nova if versao_numerica_valida(versao_nova) else None


def esta_atrasado(status: object, data_planejada: object, hoje: datetime) -> bool:
    """Um curso está atrasado quando está ON TIME mas a data planejada já passou.

    Repare que a exigência de ``status == "ON TIME"`` é de propósito e NÃO deve
    ser afrouxada por conta própria: mexer aqui reclassifica de uma vez todas as
    linhas já gravadas no Planos e Macs (decisão de negócio, não do código).
    Quem precisa decidir "o prazo deste ciclo já nasceu vencido?" usa
    ``prazo_ja_venceu``, que olha só a data.
    """
    if status != "ON TIME":
        return False
    data = parse_data(data_planejada)
    return data is not None and data.date() < hoje.date()


def prazo_ja_venceu(prazo: object, hoje: datetime) -> bool:
    """True se a data limite já passou em relação a ``hoje`` (o próprio dia vale).

    Usada no MOMENTO DA ESCRITA (R5): quando o ciclo é retroativo (fechar março
    estando em setembro), o prazo calculado — 1º dia do mês de destino — já
    passou. Gravar ``ON TIME`` numa linha com prazo vencido é mentira, e ainda
    fazia o estado oscilar entre execuções (``ON TIME`` -> ``ATRASADO`` ->
    ``ON TIME``, porque ``esta_atrasado`` deixa de enxergar a linha assim que ela
    vira ``ATRASADO``). Com o status derivando da data, rodar o mesmo ciclo
    várias vezes dá sempre o mesmo resultado.

    Aceita ``datetime`` ou texto ``dd/mm/aaaa`` (passa por ``parse_data``); o que
    não for data vira ``False`` (sem data = sem prazo vencido).
    """
    data = parse_data(prazo)
    return data is not None and data.date() < hoje.date()


# Resultados possíveis de `decidir_marcacao_planos` (definida no fim do
# arquivo). São sentinelas internas (não vão pra planilha): dizem O QUE
# escrever, não o texto do status.
MARCACAO_ON_TIME = "ON_TIME"
MARCACAO_ATRASADO_PRAZO_NOVO = "ATRASADO_PRAZO_NOVO"
MARCACAO_ATRASADO_PRAZO_ANTIGO = "ATRASADO_PRAZO_ANTIGO"


def detectar_alteracoes(
    linhas_matriz: list[tuple[str, object, object]],
) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    """Lê as linhas da Matriz - Atualização e aponta os cursos que versionaram.

    Cada item de ``linhas_matriz`` é ``(codigo, versao_atual, versao_nova)`` como
    lido diretamente das colunas C, E e F da planilha. Retorna:
    - ``versoes_matriz``: código -> versão nova normalizada.
    - ``alteracoes``: código -> {"v_antiga": ..., "v_nova": ...}, só para quem
      versionou (versão atual difere da nova).

    Gatilho válido na coluna F: **só número de versão**. O marcador
    ``VERIFICAR`` NÃO entra aqui (decisão A3, 2026-09-02 — reverte o T3):
    confirmado com quem mantém a Matriz que ``VERIFICAR`` significa "POP
    obsoletado ou renomeado", nunca uma versão nova de verdade. Uma linha
    ``VERIFICAR`` não versiona ninguém e não escreve nada na coluna de versão do
    planilha — só é sinalizada (ver ``codigos_verificar``). Qualquer outro
    texto não numérico (``'N/A'``, comentário solto) continua sendo ignorado —
    ver ``codigos_com_versao_nova_invalida``.

    **Coluna E (versão atual) vazia = linha incompleta (R4, 2026-09-18).** Antes,
    uma linha ``(COD, '', '5.0')`` entrava em ``versoes_matriz`` mas não em
    ``alteracoes``: a versão 5.0 podia ser quitada na coluna E e NINGUÉM era
    marcado como pendente — ficava registrado que o time está na
    versão nova sem que ninguém tivesse sido obrigado a fazer o treinamento, em
    silêncio. Agora a linha é ignorada dos dois lados (igual ao caso da coluna F
    vazia) e os códigos saem em ``codigos_sem_versao_atual`` pra alguém completar
    a Matriz.

    **Código repetido na Matriz: a ÚLTIMA linha vence (R3, 2026-09-18)**, pros
    dois lados. Antes, ``[(COD,'1.0','2.0'), (COD,'2.0','2.0')]`` deixava a
    versão da última linha em ``versoes_matriz`` mas mantinha o versionamento da
    primeira em ``alteracoes`` — as pessoas eram remarcadas e o rodapé saía
    errado. Linha pulada (F vazia/não numérica, E vazia) não conta como
    ocorrência: ela não sobrescreve o que uma linha válida já registrou. Ver
    ``codigos_duplicados_na_matriz`` pro aviso.
    """
    versoes_matriz: dict[str, str] = {}
    alteracoes: dict[str, dict[str, str]] = {}

    for codigo, versao_atual, versao_nova in linhas_matriz:
        if not codigo or not versao_nova:
            continue
        versao_nova_norm = normalizar_versao(versao_nova)
        if not versao_numerica_valida(versao_nova_norm):
            continue
        versao_atual_norm = normalizar_versao(versao_atual)
        if not versao_numerica_valida(versao_atual_norm):
            continue  # R4: linha incompleta/inválida não escreve nem versiona.
        codigo_str = str(codigo).strip()
        versoes_matriz[codigo_str] = versao_nova_norm

        # R9 (decisão de negócio, 2026-10-02): compara o NÚMERO, não o texto —
        # E '8' (texto) e F 8.0 (número) são a mesma versão. A Matriz de Maio/2026
        # tinha 47 linhas assim, que virariam versionamento falso.
        if not versoes_equivalentes(versao_atual_norm, versao_nova_norm):
            alteracoes[codigo_str] = {"v_antiga": versao_atual_norm, "v_nova": versao_nova_norm}
        else:
            # R3: esta ocorrência diz "sem alteração" e é a mais recente -> um
            # versionamento registrado por uma linha anterior do MESMO código cai.
            alteracoes.pop(codigo_str, None)

    return versoes_matriz, alteracoes


def codigos_sem_versao_nova(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[str]:
    """Códigos que têm um POP na coluna C mas a coluna F (versão nova) vazia.

    Essas linhas são ignoradas por ``detectar_alteracoes``. Sem um alerta, um
    "0 versionamentos" causado por alguém que esqueceu de preencher a coluna F
    fica indistinguível de um mês em que realmente nada versionou. A ordem de
    aparição é preservada e duplicatas são removidas.
    """
    faltantes: list[str] = []
    for codigo, _versao_atual, versao_nova in linhas_matriz:
        if not codigo or versao_nova:
            continue
        codigo_str = str(codigo).strip()
        if codigo_str and codigo_str not in faltantes:
            faltantes.append(codigo_str)
    return faltantes


def codigos_sem_versao_atual(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[str]:
    """Códigos com versão nova numérica na coluna F mas a coluna E (atual) VAZIA.

    São as "linhas incompletas" do R4 (2026-09-18). Desde essa correção elas são
    ignoradas por ``detectar_alteracoes`` — nada é escrito na coluna de versão do
    fonte de dados e ninguém vira pendência. Sem este aviso, o efeito seria o
    mesmo silêncio de antes, só que na direção oposta: o versionamento simplesmente
    não aconteceria e ninguém saberia por quê.

    Ordem de aparição, sem duplicatas (mesmo padrão de ``codigos_sem_versao_nova``).
    """
    faltantes: list[str] = []
    for codigo, versao_atual, versao_nova in linhas_matriz:
        if not codigo or not versao_nova:
            continue
        if not versao_numerica_valida(versao_nova):
            continue
        if normalizar_versao(versao_atual):
            continue
        codigo_str = str(codigo).strip()
        if codigo_str and codigo_str not in faltantes:
            faltantes.append(codigo_str)
    return faltantes


def codigos_com_versao_atual_invalida(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[tuple[str, str]]:
    """Linhas com F numérica e E preenchida, porém sem versão numérica válida."""
    invalidas: list[tuple[str, str]] = []
    vistos: set[str] = set()
    for codigo, versao_atual, versao_nova in linhas_matriz or []:
        codigo_str = str(codigo or "").strip()
        atual = normalizar_versao(versao_atual)
        if (
            not codigo_str
            or not atual
            or not versao_numerica_valida(versao_nova)
            or versao_numerica_valida(atual)
            or codigo_str in vistos
        ):
            continue
        vistos.add(codigo_str)
        invalidas.append((codigo_str, atual))
    return invalidas


def codigos_duplicados_na_matriz(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[str]:
    """Códigos que aparecem em mais de uma linha da Matriz (R3, 2026-09-18).

    Duplicata quase sempre é erro de preenchimento. Como ``detectar_alteracoes``
    resolve o conflito pela ÚLTIMA ocorrência, uma linha antiga esquecida em cima
    pode ser silenciosamente anulada por uma de baixo (ou vice-versa) — quem
    mantém a Matriz precisa saber.

    Conta qualquer linha com código preenchido, não importa o que tenha nas
    colunas E/F. Ordem de aparição (da 2ª ocorrência), sem duplicatas.
    """
    vistos: set[str] = set()
    duplicados: list[str] = []
    for codigo, _versao_atual, _versao_nova in linhas_matriz:
        if not codigo:
            continue
        codigo_str = str(codigo).strip()
        if not codigo_str:
            continue
        if codigo_str in vistos and codigo_str not in duplicados:
            duplicados.append(codigo_str)
        vistos.add(codigo_str)
    return duplicados


def regressoes_de_versao(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[tuple[str, str, str]]:
    """Linhas em que E e F são numéricas, mas F é menor que E.

    Retorna ``(codigo, versao_atual, versao_nova)``. Regressão não é um
    versionamento comum: seguir marcaria pessoas para uma versão mais antiga, por
    isso o chamador deve bloquear o ciclo para revisão humana.
    """
    regressoes: list[tuple[str, str, str]] = []
    for codigo, versao_atual, versao_nova in linhas_matriz or []:
        codigo_str = str(codigo or "").strip()
        atual = normalizar_versao(versao_atual)
        nova = normalizar_versao(versao_nova)
        if not codigo_str or not versao_numerica_valida(atual) or not versao_numerica_valida(nova):
            continue
        if float(nova) < float(atual):
            regressoes.append((codigo_str, atual, nova))
    return regressoes


def codigos_com_versao_nova_invalida(
    linhas_matriz: list[tuple[str, object, object]],
) -> list[tuple[str, str]]:
    """Códigos com a coluna F preenchida com algo que não é gatilho de versão.

    Ex.: ``'N/A'``, um comentário. ``detectar_alteracoes`` pula essas linhas.
    O marcador ``VERIFICAR`` NÃO entra aqui — não é "lixo", é um caso à parte
    (POP obsoletado ou renomeado; ver ``codigos_verificar``).
    Retorna pares ``(codigo, valor_bruto)``, em ordem de aparição, sem duplicatas.
    """
    fora: list[tuple[str, str]] = []
    vistos: set[str] = set()
    for codigo, _versao_atual, versao_nova in linhas_matriz:
        if not codigo or not versao_nova:
            continue
        versao_nova_norm = normalizar_versao(versao_nova)
        if versao_nova_norm == MARCADOR_VERIFICAR or versao_numerica_valida(versao_nova_norm):
            continue
        codigo_str = str(codigo).strip()
        if codigo_str and codigo_str not in vistos:
            vistos.add(codigo_str)
            fora.append((codigo_str, str(versao_nova).strip()))
    return fora


def codigos_verificar(linhas_matriz: list[tuple[str, object, object]]) -> list[str]:
    """Códigos cuja coluna F normaliza pra ``VERIFICAR`` (decisão A3, 2026-09-02).

    ``VERIFICAR`` significa "POP obsoletado ou renomeado" — nunca é versão nova
    de verdade, então NÃO entra em ``detectar_alteracoes``/``versoes_matriz``.
    Esta função é o jeito de o resto do código (aviso no terminal, linha extra
    no rodapé, `conferir`) saber quais códigos precisam de um humano olhar.

    Independe da coluna E (versão atual) — mesmo sem versão atual preenchida, o
    código entra na lista. Ordem de aparição, sem duplicatas (mesmo padrão de
    ``codigos_sem_versao_nova``).
    """
    achados: list[str] = []
    for codigo, _versao_atual, versao_nova in linhas_matriz:
        if not codigo or not versao_nova:
            continue
        if normalizar_versao(versao_nova) != MARCADOR_VERIFICAR:
            continue
        codigo_str = str(codigo).strip()
        if codigo_str and codigo_str not in achados:
            achados.append(codigo_str)
    return achados


def versoes_fora_do_padrao(
    linhas_matriz: list[tuple[str, object, object]], linha_inicial: int = 5
) -> list[str]:
    """Linhas com versão numérica (E ou F) que não é inteira, ex.: ``1.1``.

    Protege a premissa do R9 (decisão de negócio, 2026-10-02: "todos os números
    da planilha serão terminados em .0"). Como as versões são comparadas pelo
    NÚMERO, ``1.10`` e ``1.1`` seriam a mesma versão — com uma versão fracionada
    na Matriz essa comparação deixa de ser segura, então o ciclo para.
    Devolve ``"L<linha> <código> (E='1.1')"``, na ordem da planilha.
    """
    achados: list[str] = []
    for deslocamento, (codigo, atual, nova) in enumerate(linhas_matriz or []):
        codigo_str = str(codigo or "").strip()
        if not codigo_str:
            continue
        for coluna, valor in (("E", atual), ("F", nova)):
            norm = normalizar_versao(valor)
            if versao_numerica_valida(norm) and not float(norm).is_integer():
                achados.append(
                    f"L{linha_inicial + deslocamento} {codigo_str} ({coluna}='{norm}')"
                )
    return achados


def rodape_do_ciclo_ja_escrito(linhas_rodape: list[object] | None, titulo: str) -> bool:
    """True se alguma linha do rodapé é exatamente ``titulo`` (``Referente <Mês> <Ano>``)."""
    alvo = " ".join(str(titulo).split()).casefold()
    return any(
        isinstance(valor, str) and " ".join(valor.split()).casefold() == alvo
        for valor in (linhas_rodape or [])
    )


@dataclass
class AvaliacaoMatriz:
    """Resultado estruturado da validação única da Matriz.

    ``sync`` e ``conferir`` consomem este mesmo objeto. Os textos em
    ``bloqueios`` e ``avisos`` são explicações legíveis; os demais campos
    permitem que o chamador mostre detalhes sem interpretar as mensagens.
    """

    codigos_matriz: list[str] = field(default_factory=list)
    versoes_matriz: dict[str, str] = field(default_factory=dict)
    alteracoes: dict[str, dict[str, str]] = field(default_factory=dict)
    verificar: list[str] = field(default_factory=list)
    prazos: dict[str, datetime] = field(default_factory=dict)
    linhas_quitacao: dict[str, int] = field(default_factory=dict)
    bloqueios: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    sem_formula_f: list[str] = field(default_factory=list)
    e_com_formula: list[str] = field(default_factory=list)
    fora_do_padrao: list[str] = field(default_factory=list)
    linhas_rodape: list[str] = field(default_factory=list)
    sem_versao_nova: list[str] = field(default_factory=list)
    sem_versao_atual: list[str] = field(default_factory=list)
    versao_atual_invalida: list[tuple[str, str]] = field(default_factory=list)
    versao_nova_invalida: list[tuple[str, str]] = field(default_factory=list)
    duplicados: list[str] = field(default_factory=list)
    regressoes: list[tuple[str, str, str]] = field(default_factory=list)
    sem_data_aprovacao: list[str] = field(default_factory=list)
    replay: list[str] = field(default_factory=list)
    travas_abas_de_cargo: list[str] = field(default_factory=list)

    @property
    def pode_executar(self) -> bool:
        return not self.bloqueios


def versoes_equivalentes(a: object, b: object) -> bool:
    """True se ``a`` e ``b`` são a mesma versão. PURA.

    Usa a normalização que o projeto já tem, do mesmo jeito que
    ``regressoes_de_versao`` compara E com F: passa as duas por
    ``normalizar_versao`` e, quando as duas são número de versão válido
    (``versao_numerica_valida``), compara o VALOR numérico. Assim ``"8"``
    (texto na planilha, vira ``v8`` no rodapé) e ``8.0`` (número, vira
    ``v8.0``) são a mesma versão. Fora disso, compara o texto normalizado.
    """
    norm_a, norm_b = normalizar_versao(a), normalizar_versao(b)
    if versao_numerica_valida(norm_a) and versao_numerica_valida(norm_b):
        return float(norm_a) == float(norm_b)
    return norm_a == norm_b


def codigos_com_replay_no_rodape(
    alteracoes: dict[str, dict[str, str]], linhas_rodape: list[object] | None
) -> list[str]:
    """Versionamentos que já aparecem num rodapé herdado do Planos.

    Aceita tanto ``Versionado treinamento`` quanto o rótulo histórico
    ``Versionado do treinamento``. Espaços e caixa não mudam o significado.

    As versões do rodapé são comparadas por ``versoes_equivalentes``, não como
    texto: se a E era ``"8"`` (texto) no mês em que o rodapé foi escrito e é
    ``8.0`` (número) agora, ``de v8 para v9.0`` e ``de v8.0 para v9.0`` são o
    MESMO versionamento — comparar texto deixava o replay passar.
    """
    textos = [
        " ".join(str(valor).split())
        for valor in (linhas_rodape or [])
        if isinstance(valor, str) and valor.strip()
    ]
    encontrados: list[str] = []
    for codigo, mudanca in (alteracoes or {}).items():
        codigo_limpo = str(codigo).strip()
        antiga = (mudanca or {}).get("v_antiga")
        nova = (mudanca or {}).get("v_nova")
        padrao = re.compile(
            rf"^versionado\b.*(?<![\w-]){re.escape(codigo_limpo)}(?![\w-]).*"
            r"\bde\s+v(?P<antiga>\S+)\s+para\s+v(?P<nova>\S+)\s*$",
            re.IGNORECASE,
        )
        for texto in textos:
            achado = padrao.search(texto)
            if (
                achado
                and versoes_equivalentes(achado.group("antiga"), antiga)
                and versoes_equivalentes(achado.group("nova"), nova)
            ):
                encontrados.append(codigo_limpo)
                break
    return encontrados


def avaliar_matriz(
    linhas_matriz: list[tuple[str, object, object]],
    *,
    formulas_f: list[object] | None = None,
    formulas_e: list[object] | None = None,
    datas_aprovacao: dict[str, datetime] | None = None,
    linhas_rodape: list[object] | None = None,
    travas_abas_de_cargo: list[str] | None = None,
    linha_inicial: int = 5,
    dias_de_prazo: int = 30,
) -> AvaliacaoMatriz:
    """Avalia todas as regras da Matriz sem abrir ou alterar Excel.

    ``formulas_f`` segue a mesma ordem de ``linhas_matriz``. Quando é omitido,
    a validação de fórmula não é feita (conveniência para chamadas puras
    antigas); os fluxos de ``sync`` e ``conferir`` sempre a fornecem.

    ``formulas_e`` (mesma ordem) alimenta a trava "E com fórmula" (decisão do
    regra de negócio, 2026-10-02): a coluna E é a versão que as pessoas já treinaram e tem
    que ser VALOR. Com fórmula (ex.: VLOOKUP na 'Versionamento Mês') ela
    acompanha a F sozinha e o versionamento daquele POP nunca é detectado —
    Agosto/2026 tinha 8 linhas assim.

    ``travas_abas_de_cargo`` são os motivos (já em texto) pelos quais as abas de
    cargo impedem o ciclo — aba faltando, aba com dados sem bloco
    'Treinamentos'. Quem lê o workbook calcula com
    ``planos_macs.travas_das_abas_de_cargo``; aqui elas só entram nos
    ``bloqueios``, pra o ``sync`` e o ``conferir`` pararem pelo MESMO motivo.
    """
    linhas = list(linhas_matriz or [])
    resultado = AvaliacaoMatriz()
    resultado.codigos_matriz = list(
        dict.fromkeys(
            codigo_limpo
            for codigo, _versao_atual, _versao_nova in linhas
            if (codigo_limpo := str(codigo or "").strip())
        )
    )
    resultado.versoes_matriz, resultado.alteracoes = detectar_alteracoes(linhas)
    resultado.verificar = codigos_verificar(linhas)
    resultado.sem_versao_nova = codigos_sem_versao_nova(linhas)
    resultado.sem_versao_atual = codigos_sem_versao_atual(linhas)
    resultado.versao_atual_invalida = codigos_com_versao_atual_invalida(linhas)
    resultado.versao_nova_invalida = codigos_com_versao_nova_invalida(linhas)
    resultado.duplicados = codigos_duplicados_na_matriz(linhas)
    resultado.regressoes = regressoes_de_versao(linhas)

    if formulas_f is not None:
        for indice, ((codigo, _atual, nova), formula) in enumerate(
            zip(linhas, formulas_f), start=linha_inicial
        ):
            codigo_limpo = str(codigo or "").strip()
            if (
                codigo_limpo
                and str(nova or "").strip()
                and not str(formula or "").strip().startswith("=")
            ):
                resultado.sem_formula_f.append(f"L{indice} {codigo_limpo}")

    if formulas_e is not None:
        for indice, ((codigo, _atual, _nova), formula) in enumerate(
            zip(linhas, formulas_e), start=linha_inicial
        ):
            codigo_limpo = str(codigo or "").strip()
            if codigo_limpo and str(formula or "").strip().startswith("="):
                resultado.e_com_formula.append(f"L{indice} {codigo_limpo}")

    resultado.fora_do_padrao = versoes_fora_do_padrao(linhas, linha_inicial)
    resultado.linhas_rodape = [
        " ".join(valor.split())
        for valor in (linhas_rodape or [])
        if isinstance(valor, str) and valor.strip()
    ]

    # Guarda a linha da última ocorrência válida; se ela diz que não houve
    # mudança, o código também sai da quitação, igual sai de ``alteracoes``.
    for deslocamento, (codigo, atual, nova) in enumerate(linhas):
        codigo_limpo = str(codigo or "").strip()
        if (
            not codigo_limpo
            or not versao_numerica_valida(atual)
            or not versao_numerica_valida(nova)
        ):
            continue
        if not versoes_equivalentes(atual, nova):
            resultado.linhas_quitacao[codigo_limpo] = linha_inicial + deslocamento
        else:
            resultado.linhas_quitacao.pop(codigo_limpo, None)

    if datas_aprovacao is not None:
        resultado.sem_data_aprovacao = sorted(
            codigo for codigo in resultado.alteracoes if codigo not in datas_aprovacao
        )
        resultado.prazos = {
            codigo: datas_aprovacao[codigo] + timedelta(days=dias_de_prazo)
            for codigo in resultado.alteracoes
            if codigo in datas_aprovacao
        }

    resultado.replay = codigos_com_replay_no_rodape(
        resultado.alteracoes, linhas_rodape
    )

    if resultado.sem_formula_f:
        resultado.bloqueios.append(
            "F preenchida sem fórmula: " + ", ".join(resultado.sem_formula_f)
        )
    if resultado.e_com_formula:
        resultado.bloqueios.append(
            "E com fórmula (troque pelo VALOR da versão vigente, a que as pessoas "
            "já treinaram): " + ", ".join(resultado.e_com_formula)
        )
    if resultado.fora_do_padrao:
        resultado.bloqueios.append(
            "versão fora do padrão N.0 (as versões são comparadas como número): "
            + ", ".join(resultado.fora_do_padrao)
        )
    if resultado.sem_versao_atual:
        resultado.bloqueios.append(
            "E vazia com F numérica: " + ", ".join(resultado.sem_versao_atual)
        )
    if resultado.versao_atual_invalida:
        resultado.bloqueios.append(
            "E não numérica com F numérica: "
            + ", ".join(f"{c} (E='{v}')" for c, v in resultado.versao_atual_invalida)
        )
    if resultado.regressoes:
        resultado.bloqueios.append(
            "regressão de versão: "
            + ", ".join(f"{c} ({a} -> {n})" for c, a, n in resultado.regressoes)
        )
    if resultado.sem_data_aprovacao:
        resultado.bloqueios.append(
            "versionado sem data de aprovação: " + ", ".join(resultado.sem_data_aprovacao)
        )
    if resultado.replay:
        resultado.bloqueios.append(
            "versionamento já registrado no rodapé: " + ", ".join(resultado.replay)
        )
    resultado.travas_abas_de_cargo = [
        str(trava) for trava in (travas_abas_de_cargo or []) if str(trava or "").strip()
    ]
    resultado.bloqueios.extend(resultado.travas_abas_de_cargo)

    if resultado.sem_versao_nova:
        resultado.avisos.append("F vazia: " + ", ".join(resultado.sem_versao_nova))
    if resultado.versao_nova_invalida:
        resultado.avisos.append(
            "F com texto: "
            + ", ".join(f"{c} (F='{v}')" for c, v in resultado.versao_nova_invalida)
        )
    if resultado.verificar:
        resultado.avisos.append("VERIFICAR: " + ", ".join(resultado.verificar))
    if resultado.duplicados:
        resultado.avisos.append(
            "duplicata na Matriz (vale a última): " + ", ".join(resultado.duplicados)
        )
    return resultado


MARCACAO_PRESERVAR = "PRESERVAR"
MARCACAO_STATUS_DESCONHECIDO = "STATUS_DESCONHECIDO"

# Textos do Planejado que dizem que a pessoa ESCREVEU ou APROVOU o POP: com
# status OK, o versionamento não cobra o curso dela e a célula fica como está.
# Só estas 4 palavras, comparadas já normalizadas (ver
# `planejado_de_elaborador_ou_aprovador`).
PLANEJADO_ELABORADOR_OU_APROVADOR = frozenset(
    {"ELABORADOR", "ELABORADORA", "APROVADOR", "APROVADORA"}
)


def planejado_de_elaborador_ou_aprovador(planejado: object) -> bool:
    """True se o Planejado é ``ELABORADOR``/``ELABORADORA``/``APROVADOR``/``APROVADORA``.

    Compara depois de ``normalize_name`` (tira espaço das bordas, sobe pra
    CAIXA ALTA e tira acento): os dados reais têm ``'Elaborador'``,
    ``'Aprovadora'``, ``'Elaboradora'`` e ``'Elaboradora '`` (com espaço no
    fim). A regra antiga só aceitava ``ELABORADOR`` e ``APROVADORA``, então
    ``'Elaboradora'`` (caso real, abas Gerente e Analista Sr) perdia o texto
    e virava ``ON TIME``. Palavra parecida (``'Elaboração'``) NÃO conta. Só
    decide; quem preserva a célula não reescreve nada, então o texto original
    fica intacto.
    """
    if planejado is None:
        return False
    return normalize_name(planejado) in PLANEJADO_ELABORADOR_OU_APROVADOR


def normalizar_status(status: object) -> str:
    """Normalização compartilhada dos status de Planos: bordas e caixa."""
    return str(status or "").strip().upper()


def decidir_marcacao_planos(
    status: object,
    planejado: object,
    prazo_novo: object,
    hoje: datetime,
) -> str:
    """Aplica a tabela de status do Planos quando um curso versiona."""
    status_norm = normalizar_status(status)

    # Espec §2: status vazio = tratar como OK (a pessoa está no bloco, precisa
    # do curso). Normaliza ANTES da regra do ELABORADOR/APROVADORA: senão um
    # vazio com esse texto no Planejado virava ON TIME e o texto sumia.
    if not status_norm:
        status_norm = "OK"

    if status_norm == "OK" and planejado_de_elaborador_ou_aprovador(planejado):
        return MARCACAO_PRESERVAR
    if status_norm == "OK":
        return (
            MARCACAO_ATRASADO_PRAZO_NOVO
            if prazo_ja_venceu(prazo_novo, hoje)
            else MARCACAO_ON_TIME
        )
    if status_norm == "ON TIME":
        if prazo_ja_venceu(planejado, hoje):
            return MARCACAO_ATRASADO_PRAZO_ANTIGO
        return (
            MARCACAO_ATRASADO_PRAZO_NOVO
            if prazo_ja_venceu(prazo_novo, hoje)
            else MARCACAO_ON_TIME
        )
    if status_norm == "ATRASADO":
        return MARCACAO_ATRASADO_PRAZO_ANTIGO
    if status_norm in {"NA", "N/A", "HSE", "OUTROS", "GLOBAL"}:
        return MARCACAO_PRESERVAR
    return MARCACAO_STATUS_DESCONHECIDO
