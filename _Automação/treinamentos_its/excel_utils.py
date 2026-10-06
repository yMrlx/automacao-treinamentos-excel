"""Funções pequenas e sem efeitos colaterais para tratar dados do Excel."""

import unicodedata
from datetime import datetime

from treinamentos_its.config import MARCADOR_VERIFICAR


def sem_acentos(texto: str) -> str:
    """Tira os acentos de um texto (NFKD + descarta os caracteres combinantes).

    ``'João'`` -> ``'Joao'``, ``'Estagiário'`` -> ``'Estagiario'``. Não mexe em
    caixa nem em espaços — isso é com quem chama.
    """
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def normalize_name(nome: object) -> str:
    """Normaliza um nome para comparações entre planilhas.

    Tira espaços das bordas, sobe pra CAIXA ALTA e **remove acentos**, para que
    variações de digitação como "João" e "Joao" sejam a mesma pessoa.
    """
    return sem_acentos(str(nome).strip().upper())


def normalizar_versao(versao: object) -> str:
    """Converte versões numéricas e textuais para uma representação comparável.

    O marcador ``VERIFICAR`` (ver ``config.MARCADOR_VERIFICAR``) é padronizado em
    CAIXA ALTA só pra ser RECONHECIDO sem depender de caixa/espaço. Ele NÃO é
    versionamento: na coluna F da Matriz significa "POP obsoletado ou
    renomeado" (decisão A3) — não versiona ninguém, só gera aviso e a linha
    própria no rodapé (ver ``dominio.codigos_verificar``).
    """
    if versao is None:
        return ""
    try:
        texto = str(float(versao)) if isinstance(versao, (int, float)) else str(versao).strip()
    except (TypeError, ValueError):
        texto = str(versao).strip()
    # Apóstrofo da frente = "isto é texto" no Excel (ver `versao_para_gravar`). No
    # Excel real ele nem chega no valor; aqui é por garantia, como no `parse_data`.
    texto = texto.lstrip("'").strip()
    return MARCADOR_VERIFICAR if texto.upper() == MARCADOR_VERIFICAR else texto


def versao_e_numerica(versao_normalizada: str) -> bool:
    """True se a versão representa um número (ex.: ``'9.0'``, ``'16'``).

    Usado para descartar textos que não são versão de verdade na coluna F da
    Matriz — ``'VERIFICAR'``, ``'N/A'``, ``'-'``, célula com comentário solto.
    """
    if not versao_normalizada:
        return False
    try:
        float(versao_normalizada)
        return True
    except (TypeError, ValueError):
        return False


def normalizar_cabecalho(valor: object) -> str:
    """Título de coluna comparável: sem espaço nas bordas, sem acento, minúsculo.

    Morava no ``conferir``; subiu pra cá quando o R12 passou a precisar da mesma
    coisa pra achar a coluna da data de aprovação na aba 'Versionamento Mês'
    (``conferir`` importa ``sync``, então ``sync`` não pode importar ``conferir``).
    """
    return sem_acentos(str(valor or "").strip().casefold())


def localizar_coluna(cabecalho: list | None, nomes_aceitos: tuple[str, ...]) -> int | None:
    """Índice 1-based da coluna cujo título bate com um dos ``nomes_aceitos``.

    Pura (sem Excel). A ordem de ``nomes_aceitos`` é a ordem de PREFERÊNCIA: o
    primeiro nome é procurado na linha inteira antes do segundo ("Major Version
    Number" ganha de "Version" no export de obsoletos). Dentro do mesmo nome vale
    a coluna mais à esquerda — o export de Julho/2026 tem duas colunas "Version"
    e é a primeira (``9.0``) que casa com o formato da Matriz.

    Devolve ``None`` quando nenhum bate — o chamador transforma isso em erro
    explícito em vez de adivinhar um índice. Procurar pelo NOME em vez de por
    índice fixo não é preciosismo: o layout desses exports muda de um mês pro
    outro (ver ``conferir`` e ``versionamento``).
    """
    titulos = [normalizar_cabecalho(celula) for celula in (cabecalho or [])]
    for nome in nomes_aceitos:
        alvo = normalizar_cabecalho(nome)
        for indice, titulo in enumerate(titulos, start=1):
            if titulo == alvo:
                return indice
    return None


def parse_data(valor: object) -> datetime | None:
    """Lê datas retornadas por Excel nos formatos utilizados nas planilhas."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor
    if isinstance(valor, str):
        data_texto = valor.replace("'", "").strip()
        # Ordem importa: dd/mm/aaaa (padrão das planilhas) primeiro; aaaa-mm-dd
        # (ISO) em seguida; mm/dd/aaaa só entra para datas em que o dia > 12
        # torna a leitura inequívoca.
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(data_texto, formato)
            except ValueError:
                continue
    return None


def versao_como_texto(valor: object) -> object:
    """Versão numérica inteira vira o texto ``'N.0'``; o resto fica como veio. PURA.

    O export traz a versão como TEXTO (``'1.0'``) e todos os meses colados à mão
    ficaram assim — a Matriz compara E com F e pinta de vermelho quando diferem.
    Em Maio/2026 o export veio com NÚMERO e a Matriz inteira ficou vermelha
    (E ``'6.0'`` x F ``6``). Normalizar pro texto do resto do histórico evita
    isso; versão fracionada (1.1) fica como veio e a trava do ciclo pega.
    """
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return valor
    if float(valor).is_integer():
        return f"{int(valor)}.0"
    return valor


def versao_para_gravar(valor: object) -> object:
    """Versão pronta pra ``Range.Value``: TEXTO protegido por apóstrofo.

    Gravar ``'4.0'`` direto pelo COM faz o Excel converter pra NÚMERO 4 — foi o
    que a quitação E=F fez no Setembro/2026 (E25 = 4, linha vermelha na Matriz e
    a linha "Versão" dos blocos mostrando ``4``). Mesmo truque de
    ``formatar_data_texto``: o apóstrofo some do valor e não aparece na tela.
    """
    texto = versao_como_texto(valor)
    return f"'{texto}" if isinstance(texto, str) and texto else texto


def formatar_data_texto(planejado: object) -> str:
    """Formata a data como texto para impedir alterações automáticas do Excel."""
    if hasattr(planejado, "day"):
        return f"'{planejado.day:02d}/{planejado.month:02d}/{planejado.year}"
    data = parse_data(planejado)
    if data:
        return f"'{data.day:02d}/{data.month:02d}/{data.year}"
    return f"'{planejado}"


def limpar_celula_status(celula) -> None:
    """Esvazia uma célula de status/data e tira a formatação (fundo + fonte).

    "ON TIME"/"ATRASADO" e a data planejada são sinalizadores TEMPORÁRIOS de
    pendência dentro do prazo, não um estado gravado. Quando a pessoa conclui o
    curso (decisão de negócio de 2026-09-02), a pendência simplesmente some: a
    célula volta a ficar vazia e sem cor. O `sync` só reintroduz "ON TIME"
    quando o curso versiona de novo.
    """
    celula.value = None
    celula.color = None
    celula.font.color = (0, 0, 0)
