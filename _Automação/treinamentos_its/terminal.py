"""Perguntas sim/não no terminal, à prova de execução não-interativa."""

from treinamentos_its.logger import log


def perguntar_sim_nao(pergunta: str, *, assumir_em_eof: bool = False) -> bool:
    """Lê ``s``/``n`` do terminal.

    Sem entrada interativa (``EOFError`` — pipe, agendador, duplo-clique sem
    console) a função devolve ``assumir_em_eof``. O padrão é ``False`` (= ``n``),
    que é o lado seguro na dúvida: não faz nada.
    """
    try:
        resposta = input(pergunta).strip().lower()
    except EOFError:
        log.warning(
            "Sem entrada interativa disponível; assumindo '%s'.",
            "s" if assumir_em_eof else "n",
        )
        return assumir_em_eof
    return resposta in ("s", "sim")
