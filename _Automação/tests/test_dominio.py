from datetime import datetime

import pytest

from treinamentos_its.dominio import (
    codigos_com_versao_nova_invalida,
    codigos_duplicados_na_matriz,
    codigos_com_versao_atual_invalida,
    codigos_sem_versao_atual,
    codigos_sem_versao_nova,
    codigos_verificar,
    detectar_alteracoes,
    esta_atrasado,
    prazo_ja_venceu,
    regressoes_de_versao,
    valor_para_quitacao,
    versao_numerica_valida,
)


# --- Planos e Macs como fonte única (decisão de 25/09) --------------------


@pytest.mark.parametrize(
    ("status", "planejado", "prazo", "esperado"),
    [
        ("OK", None, datetime(2026, 10, 30), "ON_TIME"),
        ("OK", None, datetime(2026, 9, 1), "ATRASADO_PRAZO_NOVO"),
        ("OK", "ELABORADOR", datetime(2026, 10, 30), "PRESERVAR"),
        ("OK", "APROVADORA", datetime(2026, 10, 30), "PRESERVAR"),
        ("ON TIME", datetime(2026, 9, 1), datetime(2026, 10, 30), "ATRASADO_PRAZO_ANTIGO"),
        ("ON TIME", datetime(2026, 10, 1), datetime(2026, 10, 30), "ON_TIME"),
        ("ON TIME", datetime(2026, 10, 1), datetime(2026, 9, 1), "ATRASADO_PRAZO_NOVO"),
        ("ATRASADO", datetime(2026, 8, 1), datetime(2026, 10, 30), "ATRASADO_PRAZO_ANTIGO"),
        ("NA", None, datetime(2026, 10, 30), "PRESERVAR"),
        ("N/A", None, datetime(2026, 10, 30), "PRESERVAR"),
        ("HSE", None, datetime(2026, 10, 30), "PRESERVAR"),
        ("OUTROS", None, datetime(2026, 10, 30), "PRESERVAR"),
        ("GLOBAL", None, datetime(2026, 10, 30), "PRESERVAR"),
        (None, None, datetime(2026, 10, 30), "ON_TIME"),
        ("TEXTO NOVO", None, datetime(2026, 10, 30), "STATUS_DESCONHECIDO"),
    ],
)
def test_tabela_completa_de_status_do_planos(status, planejado, prazo, esperado):
    from treinamentos_its.dominio import decidir_marcacao_planos

    assert decidir_marcacao_planos(
        status, planejado, prazo, datetime(2026, 9, 25)
    ) == esperado


def test_avaliar_matriz_reune_todas_as_travas_e_avisos():
    from treinamentos_its.dominio import avaliar_matriz

    linhas = [
        ("SEM-FORMULA", "1.0", "2.0"),
        ("SEM-E", None, "2.0"),
        ("E-TEXTO", "x", "2.0"),
        ("REGRESSAO", "3.0", "2.0"),
        ("SEM-DATA", "1.0", "2.0"),
        ("VAZIA", "1.0", None),
        ("TEXTO", "1.0", "comentário"),
        ("VER", "1.0", "VERIFICAR"),
        ("DUP", "1.0", "2.0"),
        ("DUP", "2.0", "3.0"),
    ]
    formulas = [None] + ["=PROCV()"] * (len(linhas) - 1)
    resultado = avaliar_matriz(
        linhas,
        formulas_f=formulas,
        datas_aprovacao={"DUP": datetime(2026, 9, 1)},
        linhas_rodape=["Versionado treinamento DUP de v2.0 para v3.0"],
    )

    assert not resultado.pode_executar
    assert resultado.sem_formula_f == ["L5 SEM-FORMULA"]
    assert resultado.sem_versao_atual == ["SEM-E"]
    assert resultado.versao_atual_invalida == [("E-TEXTO", "x")]
    assert resultado.regressoes == [("REGRESSAO", "3.0", "2.0")]
    assert "SEM-DATA" in resultado.sem_data_aprovacao
    assert resultado.replay == ["DUP"]
    assert resultado.verificar == ["VER"]
    assert resultado.duplicados == ["DUP"]


def test_avaliar_matriz_valida_calcula_prazo_e_linha_de_quitacao():
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz(
        [("POP-1", "1.0", "2.0")],
        formulas_f=["=PROCV()"],
        datas_aprovacao={"POP-1": datetime(2026, 9, 1)},
    )

    assert resultado.pode_executar
    assert resultado.prazos == {"POP-1": datetime(2026, 10, 1)}
    assert resultado.linhas_quitacao == {"POP-1": 5}


def test_avaliar_matriz_preserva_todos_os_codigos_da_coluna_c():
    resultado = avaliar_matriz(
        [
            (" NUMERICO ", "1.0", "2.0"),
            ("F-VAZIA", "1.0", None),
            ("F-TEXTO", "1.0", "comentário"),
            ("NUMERICO", "1.0", "2.0"),
            (None, "1.0", "2.0"),
        ],
        formulas_f=["=PROCV()"] * 5,
        datas_aprovacao={"NUMERICO": datetime(2026, 9, 1)},
    )

    assert resultado.codigos_matriz == ["NUMERICO", "F-VAZIA", "F-TEXTO"]

HOJE = datetime(2026, 8, 10)


def test_quitacao_aceita_apenas_versao_numerica_segura():
    assert valor_para_quitacao(9.0) == 9.0
    assert valor_para_quitacao(" 10.5 ") == " 10.5 "
    assert valor_para_quitacao("VERIFICAR") is None
    assert valor_para_quitacao("N/A") is None
    assert valor_para_quitacao(None) is None


def test_versao_numerica_rejeita_codigos_de_erro_e_nao_finitos():
    assert versao_numerica_valida(-2146826246) is False
    assert versao_numerica_valida(float("nan")) is False
    assert versao_numerica_valida(float("inf")) is False
    assert versao_numerica_valida(True) is False


def test_versao_atual_invalida_e_listada_quando_f_e_numerica():
    linhas = [("POP-1", "VERIFICAR", "4.0"), ("POP-2", "3.0", "4.0")]
    assert codigos_com_versao_atual_invalida(linhas) == [("POP-1", "VERIFICAR")]


def test_regressao_de_versao_e_detectada_e_textos_sao_ignorados():
    linhas = [
        ("POP-1", "9.0", "8.0"),
        ("POP-2", "2.0", "3.0"),
        ("POP-3", "4.0", "VERIFICAR"),
    ]
    assert regressoes_de_versao(linhas) == [("POP-1", "9.0", "8.0")]


def test_esta_atrasado_falso_quando_status_nao_e_on_time():
    assert esta_atrasado("ATRASADO", "01/01/2026", HOJE) is False
    assert esta_atrasado(None, "01/01/2026", HOJE) is False


def test_esta_atrasado_true_quando_data_planejada_ja_passou():
    assert esta_atrasado("ON TIME", "01/01/2026", HOJE) is True


def test_esta_atrasado_false_quando_data_planejada_e_futura():
    assert esta_atrasado("ON TIME", "01/01/2027", HOJE) is False


def test_esta_atrasado_false_quando_sem_data_planejada():
    assert esta_atrasado("ON TIME", None, HOJE) is False


def test_esta_atrasado_false_quando_status_tem_espaco_sobrando():
    # DOC/BUG LATENTE: a comparação é "== 'ON TIME'" exata. "ON TIME " (espaço no
    # fim), que o Excel às vezes devolve, NÃO conta como atrasado - a linha
    # escapa da marcação. (`esta_atrasado` não normaliza o status.)
    assert esta_atrasado("ON TIME ", "01/01/2026", HOJE) is False
    assert esta_atrasado(" ON TIME", "01/01/2026", HOJE) is False


def test_esta_atrasado_no_limite_do_dia_de_hoje_nao_e_atraso():
    # data planejada == hoje: ainda dentro do prazo (usa "<", não "<=").
    assert esta_atrasado("ON TIME", "10/08/2026", HOJE) is False


def test_detectar_alteracoes_sem_mudanca_de_versao():
    versoes, alteracoes = detectar_alteracoes([("COD1", "1.0", "1.0")])
    assert versoes == {"COD1": "1.0"}
    assert alteracoes == {}


def test_detectar_alteracoes_com_versionamento():
    versoes, alteracoes = detectar_alteracoes([("COD1", "1.0", "2.0")])
    assert versoes == {"COD1": "2.0"}
    assert alteracoes == {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_ignora_linha_sem_codigo_ou_sem_versao_nova():
    versoes, alteracoes = detectar_alteracoes([(None, "1.0", "2.0"), ("COD1", "1.0", None)])
    assert versoes == {}
    assert alteracoes == {}


def test_detectar_alteracoes_sem_versao_atual_ignora_a_linha_inteira():
    # R4 (2026-09-18) - MUDANÇA DE COMPORTAMENTO: antes este teste exigia
    # `versoes == {"COD1": "3.0"}`, ou seja, a versão nova era escrita na coluna E
    # do Procedimentos sem ninguém virar pendência. Linha com coluna E vazia agora
    # é "linha incompleta": não escreve versão e não versiona ninguém.
    versoes, alteracoes = detectar_alteracoes([("COD1", None, "3.0")])
    assert versoes == {}
    assert alteracoes == {}


def test_detectar_alteracoes_coluna_e_vazia_nao_grava_versao_nem_pendencia():
    # O cenário exato do R4, reproduzido na revisão: ("COD2", "", "5.0") devolvia
    # ({'COD2': '5.0'}, {}) - versão registrada, ninguém obrigado a treinar.
    versoes, alteracoes = detectar_alteracoes([("COD2", "", "5.0")])
    assert versoes == {}
    assert alteracoes == {}


def test_detectar_alteracoes_linha_incompleta_nao_atrapalha_as_outras():
    versoes, alteracoes = detectar_alteracoes([("COD2", "", "5.0"), ("COD3", "1.0", "2.0")])
    assert versoes == {"COD3": "2.0"}
    assert alteracoes == {"COD3": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_ignora_versao_nova_nao_numerica():
    # Texto solto na coluna F (que NÃO seja "VERIFICAR") -> linha ignorada por completo.
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "3.0", "N/A"), ("COD2", "1.0", "2.0")]
    )
    assert versoes == {"COD2": "2.0"}
    assert alteracoes == {"COD2": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_verificar_nao_versiona_ninguem():
    # Decisão A3 (2026-09-02, reverte o T3): "VERIFICAR" = POP obsoletado ou
    # renomeado, NUNCA é versão nova de verdade. Não entra em versoes_matriz nem
    # em alteracoes - só o código numérico ao lado segue seu fluxo normal.
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "3.0", "VERIFICAR"), ("COD2", "1.0", "2.0")]
    )
    assert versoes == {"COD2": "2.0"}
    assert alteracoes == {"COD2": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_verificar_caixa_e_espaco_tambem_nao_versiona():
    versoes, alteracoes = detectar_alteracoes([("COD1", "3.0", "  verificar ")])
    assert versoes == {}
    assert alteracoes == {}


def test_detectar_alteracoes_verificar_sem_versao_atual_tambem_nao_versiona():
    versoes, alteracoes = detectar_alteracoes([("COD1", None, "VERIFICAR")])
    assert versoes == {}
    assert alteracoes == {}


def test_detectar_alteracoes_aceita_versao_numerica_inteira():
    versoes, alteracoes = detectar_alteracoes([("COD1", 8, 9)])
    assert versoes == {"COD1": "9.0"}
    assert alteracoes == {"COD1": {"v_antiga": "8.0", "v_nova": "9.0"}}


def test_detectar_alteracoes_codigo_duplicado_na_matriz_ultima_linha_vence():
    # Código repetido na Matriz (dado sujo): a última ocorrência sobrescreve.
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "1.0", "2.0"), ("COD1", "1.0", "3.0")]
    )
    assert versoes == {"COD1": "3.0"}
    assert alteracoes == {"COD1": {"v_antiga": "1.0", "v_nova": "3.0"}}


def test_detectar_alteracoes_codigo_duplicado_com_2a_linha_invalida_mantem_a_1a():
    # 1ª ocorrência válida, 2ª com texto solto -> a 2ª é pulada por completo
    # (não zera o que a 1ª gravou).
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "1.0", "2.0"), ("COD1", "1.0", "N/A")]
    )
    assert versoes == {"COD1": "2.0"}
    assert alteracoes == {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_duplicado_com_ultima_linha_sem_alteracao_derruba_o_versionamento():
    # Cenário exato do R3: a 1ª linha diz 1.0 -> 2.0 e a 2ª (mais recente) diz
    # "já está em 2.0". Antes, `versoes_matriz` usava a última linha e
    # `alteracoes` continuava com o versionamento da primeira - as pessoas eram
    # remarcadas à toa e o rodapé saía errado.
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "1.0", "2.0"), ("COD1", "2.0", "2.0")]
    )
    assert versoes == {"COD1": "2.0"}
    assert alteracoes == {}


def test_detectar_alteracoes_duplicado_linha_incompleta_nao_apaga_a_anterior():
    # A 2ª linha é pulada por completo (coluna E vazia, R4), então não "vence"
    # nada: o versionamento da 1ª continua valendo.
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "1.0", "2.0"), ("COD1", "", "2.0")]
    )
    assert versoes == {"COD1": "2.0"}
    assert alteracoes == {"COD1": {"v_antiga": "1.0", "v_nova": "2.0"}}


def test_detectar_alteracoes_duplicado_com_tres_linhas_vale_a_ultima():
    versoes, alteracoes = detectar_alteracoes(
        [("COD1", "1.0", "2.0"), ("COD1", "2.0", "2.0"), ("COD1", "2.0", "4.0")]
    )
    assert versoes == {"COD1": "4.0"}
    assert alteracoes == {"COD1": {"v_antiga": "2.0", "v_nova": "4.0"}}


# --- codigos_sem_versao_atual (R4, 2026-09-18) -------------------------------


def test_codigos_sem_versao_atual_lista_linhas_incompletas():
    linhas = [("COD1", "1.0", "2.0"), ("COD2", "", "5.0"), ("COD3", None, 4)]
    assert codigos_sem_versao_atual(linhas) == ["COD2", "COD3"]


def test_codigos_sem_versao_atual_ignora_quem_nao_tem_gatilho_valido():
    # Sem coluna F, com texto solto ou com VERIFICAR não é "linha incompleta":
    # esses casos já têm avisos próprios.
    linhas = [("COD1", None, None), ("COD2", None, "N/A"), ("COD3", None, "VERIFICAR")]
    assert codigos_sem_versao_atual(linhas) == []


def test_codigos_sem_versao_atual_remove_duplicatas_preservando_ordem():
    linhas = [("COD2", None, "1.0"), ("COD1", None, "1.0"), ("COD2", "", "3.0")]
    assert codigos_sem_versao_atual(linhas) == ["COD2", "COD1"]


# --- codigos_duplicados_na_matriz (R3, 2026-09-18) ---------------------------


def test_codigos_duplicados_na_matriz_aponta_repetido():
    linhas = [("COD1", "1.0", "2.0"), ("COD2", "1.0", "1.0"), ("COD1", "2.0", "2.0")]
    assert codigos_duplicados_na_matriz(linhas) == ["COD1"]


def test_codigos_duplicados_na_matriz_conta_qualquer_linha_com_codigo():
    # Duplicata é erro de preenchimento mesmo quando as colunas E/F estão vazias.
    linhas = [("COD1", None, None), ("COD1", None, None)]
    assert codigos_duplicados_na_matriz(linhas) == ["COD1"]


def test_codigos_duplicados_na_matriz_sem_repeticao():
    assert codigos_duplicados_na_matriz([("COD1", "1.0", "2.0"), ("COD2", "1.0", "2.0")]) == []


def test_codigos_duplicados_na_matriz_aparece_uma_vez_so_na_lista():
    linhas = [("COD1", "1.0", "2.0")] * 4
    assert codigos_duplicados_na_matriz(linhas) == ["COD1"]


def test_codigos_duplicados_na_matriz_ignora_linha_sem_codigo():
    assert codigos_duplicados_na_matriz([(None, "1.0", "2.0"), ("", "1.0", "2.0")]) == []


# --- prazo_ja_venceu (R5, 2026-09-18) ----------------------------------------


def test_prazo_ja_venceu_com_data_no_passado():
    assert prazo_ja_venceu(datetime(2026, 3, 1), HOJE) is True


def test_prazo_ja_venceu_false_no_proprio_dia_e_no_futuro():
    assert prazo_ja_venceu(datetime(2026, 8, 10), HOJE) is False
    assert prazo_ja_venceu(datetime(2026, 12, 1), HOJE) is False


def test_prazo_ja_venceu_aceita_texto_dd_mm_aaaa():
    assert prazo_ja_venceu("01/03/2026", HOJE) is True
    assert prazo_ja_venceu("01/12/2026", HOJE) is False


def test_prazo_ja_venceu_sem_data_e_false():
    assert prazo_ja_venceu(None, HOJE) is False
    assert prazo_ja_venceu("", HOJE) is False
    assert prazo_ja_venceu("qualquer coisa", HOJE) is False


def test_prazo_ja_venceu_nao_depende_do_status():
    # A diferença pra `esta_atrasado`: aqui não existe status nenhum na conta.
    # É o que permite decidir "esta linha nasce ATRASADA" sem mexer na semântica
    # global de `esta_atrasado` (que exige status == "ON TIME").
    assert prazo_ja_venceu(datetime(2026, 1, 1), HOJE) is True
    assert esta_atrasado("ATRASADO", datetime(2026, 1, 1), HOJE) is False


def test_codigos_com_versao_nova_invalida_dedup_codigo_repetido():
    linhas = [("COD1", "1.0", "N/A"), ("COD1", "1.0", "xpto")]
    assert codigos_com_versao_nova_invalida(linhas) == [("COD1", "N/A")]


def test_codigos_sem_versao_nova_lista_codigos_com_coluna_f_vazia():
    linhas = [("COD1", "1.0", "2.0"), ("COD2", "1.0", None), ("COD3", "2.0", "")]
    assert codigos_sem_versao_nova(linhas) == ["COD2", "COD3"]


def test_codigos_sem_versao_nova_ignora_linhas_sem_codigo():
    assert codigos_sem_versao_nova([(None, "1.0", None), ("", None, None)]) == []


def test_codigos_sem_versao_nova_remove_duplicatas_preservando_ordem():
    linhas = [("COD2", None, None), ("COD1", None, None), ("COD2", None, None)]
    assert codigos_sem_versao_nova(linhas) == ["COD2", "COD1"]


def test_codigos_com_versao_nova_invalida_lista_pares_codigo_valor():
    linhas = [
        ("COD1", "3.0", "VERIFICAR"),  # caso à parte (não versiona) -> NÃO é "inválido"
        ("COD2", "1.0", "2.0"),
        ("COD3", "5.0", "N/A"),
        ("COD4", "1.0", None),
    ]
    assert codigos_com_versao_nova_invalida(linhas) == [("COD3", "N/A")]


def test_codigos_com_versao_nova_invalida_vazio_quando_tudo_numerico():
    assert codigos_com_versao_nova_invalida([("COD1", "1.0", "2.0"), ("COD2", 3, 4)]) == []


# --- codigos_verificar (A3, 2026-09-02) --------------------------------------


def test_codigos_verificar_lista_codigos_com_f_verificar():
    linhas = [("COD1", "3.0", "VERIFICAR"), ("COD2", "1.0", "2.0"), ("COD3", None, "verificar")]
    assert codigos_verificar(linhas) == ["COD1", "COD3"]


def test_codigos_verificar_independe_da_coluna_e():
    # Sem versão atual preenchida, o código ainda entra na lista.
    assert codigos_verificar([("COD1", None, "VERIFICAR")]) == ["COD1"]


def test_codigos_verificar_ignora_linha_sem_codigo_ou_sem_versao_nova():
    assert codigos_verificar([(None, "1.0", "VERIFICAR"), ("COD1", "1.0", None)]) == []


def test_codigos_verificar_ignora_versao_numerica():
    assert codigos_verificar([("COD1", "1.0", "2.0")]) == []


def test_codigos_verificar_remove_duplicatas_preservando_ordem():
    linhas = [("COD2", None, "VERIFICAR"), ("COD1", None, "VERIFICAR"), ("COD2", None, "Verificar")]
    assert codigos_verificar(linhas) == ["COD2", "COD1"]


# =============================================================================
# Cobertura recriada depois do a5dad36: convergência de `decidir_marcacao_planos`
# (a função em uso) e testes das correções D e I do 510a997.
# =============================================================================

from treinamentos_its.dominio import (  # noqa: E402
    MARCACAO_ATRASADO_PRAZO_ANTIGO,
    MARCACAO_ATRASADO_PRAZO_NOVO,
    MARCACAO_ON_TIME,
    MARCACAO_PRESERVAR,
    avaliar_matriz,
    codigos_com_replay_no_rodape,
    decidir_marcacao_planos,
    versoes_equivalentes,
)

HOJE_CICLO = datetime(2026, 9, 25)
PRAZO_FUTURO = datetime(2026, 10, 30)
PRAZO_VENCIDO = datetime(2026, 9, 1)
DATA_ANTIGA = datetime(2026, 8, 1)
ON_TIME_VENCIDO = datetime(2026, 9, 10)
ON_TIME_NO_PRAZO = datetime(2026, 9, 30)


def _aplicar(celula: tuple, decisao: str, prazo) -> tuple:
    """Espelha o que `planos_macs._atualizar_aba` grava: ``(planejado, status)``."""
    planejado, _status = celula
    if decisao == MARCACAO_ON_TIME:
        return (prazo, "ON TIME")
    if decisao == MARCACAO_ATRASADO_PRAZO_NOVO:
        return (prazo, "ATRASADO")
    if decisao == MARCACAO_ATRASADO_PRAZO_ANTIGO:
        return (planejado, "ATRASADO")
    return celula  # PRESERVAR / STATUS_DESCONHECIDO: não toca


def _mesmo_ciclo_tres_vezes(celula: tuple, prazo) -> list[tuple]:
    estados = []
    for _vez in range(3):
        planejado, status = celula
        decisao = decidir_marcacao_planos(status, planejado, prazo, HOJE_CICLO)
        celula = _aplicar(celula, decisao, prazo)
        estados.append(celula)
    return estados


@pytest.mark.parametrize(
    ("inicial", "prazo", "esperado"),
    [
        # tabela §2.2 da espec, linha a linha
        ((None, "OK"), PRAZO_FUTURO, (PRAZO_FUTURO, "ON TIME")),
        ((None, "OK"), PRAZO_VENCIDO, (PRAZO_VENCIDO, "ATRASADO")),
        (("ELABORADOR", "OK"), PRAZO_FUTURO, ("ELABORADOR", "OK")),
        (("APROVADORA", "OK"), PRAZO_VENCIDO, ("APROVADORA", "OK")),
        ((ON_TIME_VENCIDO, "ON TIME"), PRAZO_FUTURO, (ON_TIME_VENCIDO, "ATRASADO")),
        ((ON_TIME_NO_PRAZO, "ON TIME"), PRAZO_FUTURO, (PRAZO_FUTURO, "ON TIME")),
        ((ON_TIME_NO_PRAZO, "ON TIME"), PRAZO_VENCIDO, (PRAZO_VENCIDO, "ATRASADO")),
        ((DATA_ANTIGA, "ATRASADO"), PRAZO_FUTURO, (DATA_ANTIGA, "ATRASADO")),
        ((DATA_ANTIGA, "ATRASADO"), PRAZO_VENCIDO, (DATA_ANTIGA, "ATRASADO")),
        ((None, "NA"), PRAZO_FUTURO, (None, "NA")),
        ((None, "N/A"), PRAZO_VENCIDO, (None, "N/A")),
        ((None, "HSE"), PRAZO_FUTURO, (None, "HSE")),
        ((None, "OUTROS"), PRAZO_FUTURO, (None, "OUTROS")),
        ((None, "GLOBAL"), PRAZO_VENCIDO, (None, "GLOBAL")),
        ((None, None), PRAZO_FUTURO, (PRAZO_FUTURO, "ON TIME")),
        ((None, None), PRAZO_VENCIDO, (PRAZO_VENCIDO, "ATRASADO")),
        (("ELABORADOR", None), PRAZO_FUTURO, ("ELABORADOR", None)),
        ((None, "REVISAR"), PRAZO_FUTURO, (None, "REVISAR")),
        # o planejado vem da planilha como texto com apóstrofo
        (("'10/09/2026", "ON TIME"), PRAZO_FUTURO, ("'10/09/2026", "ATRASADO")),
    ],
    ids=[
        "ok-futuro", "ok-vencido", "ok-elaborador", "ok-aprovadora",
        "on-time-vencido", "on-time-no-prazo", "on-time-prazo-novo-vencido",
        "atrasado-futuro", "atrasado-vencido", "na", "n/a", "hse", "outros",
        "global", "vazio-futuro", "vazio-vencido", "vazio-elaborador",
        "texto-desconhecido", "on-time-vencido-em-texto",
    ],
)
def test_mesmo_ciclo_tres_vezes_da_sempre_o_mesmo_estado(inicial, prazo, esperado):
    assert _mesmo_ciclo_tres_vezes(inicial, prazo) == [esperado] * 3


@pytest.mark.parametrize("status", [None, "", "   "])
@pytest.mark.parametrize("planejado", ["ELABORADOR", "APROVADORA", "  elaborador "])
def test_status_vazio_com_elaborador_ou_aprovadora_nao_mexe(status, planejado):
    # Correção D do 510a997: vazio vira OK ANTES da regra do ELABORADOR. Antes
    # da correção, vazio + ELABORADOR virava ON TIME e o texto sumia.
    assert decidir_marcacao_planos(status, planejado, PRAZO_FUTURO, HOJE_CICLO) == (
        MARCACAO_PRESERVAR
    )


# --- I: replay compara versões equivalentes ("8" == 8.0) --------------------------


@pytest.mark.parametrize(
    ("a", "b", "iguais"),
    [
        ("8", 8.0, True),
        ("8", "8.0", True),
        (8, "8", True),
        (" 9.0 ", 9, True),
        ("8", "8.1", False),
        ("8", "9.0", False),
        ("VERIFICAR", "verificar", True),
    ],
)
def test_versoes_equivalentes(a, b, iguais):
    assert versoes_equivalentes(a, b) is iguais


@pytest.mark.parametrize(
    ("antiga", "nova", "linha_do_rodape"),
    [
        ("8", "9.0", "Versionado treinamento DOC-POP-1 de v8.0 para v9.0"),
        ("8.0", "9.0", "Versionado treinamento DOC-POP-1 de v8 para v9.0"),
        ("8.0", "9", "Versionado treinamento DOC-POP-1 de v8.0 para v9.0"),
        ("8.0", "9.0", "  versionado do treinamento   DOC-POP-1 de v8 para v9  "),
    ],
)
def test_replay_pelo_rodape_reconhece_8_e_8_0(antiga, nova, linha_do_rodape):
    alteracoes = {"DOC-POP-1": {"v_antiga": antiga, "v_nova": nova}}
    assert codigos_com_replay_no_rodape(alteracoes, [linha_do_rodape]) == ["DOC-POP-1"]


def test_replay_nao_confunde_versionamento_diferente():
    alteracoes = {"DOC-POP-1": {"v_antiga": "8.0", "v_nova": "10.0"}}
    rodape = [
        "Versionado treinamento DOC-POP-1 de v8.0 para v9.0",
        "Versionado treinamento DOC-POP-10 de v8.0 para v10.0",
        "DOC-POP-1 — POP obsoletado ou renomeado, para um humano verificar",
        None,
        42,
    ]
    assert codigos_com_replay_no_rodape(alteracoes, rodape) == []


@pytest.mark.parametrize(
    ("e", "linha_do_rodape"),
    [
        ("8", "Versionado treinamento DOC-POP-1 de v8.0 para v9.0"),
        (8.0, "Versionado treinamento DOC-POP-1 de v8 para v9.0"),
    ],
    ids=["E em texto, rodape em numero", "E em numero, rodape em texto"],
)
def test_replay_8_e_8_0_bloqueia_o_ciclo(e, linha_do_rodape):
    resultado = avaliar_matriz(
        [("DOC-POP-1", e, 9.0)],
        formulas_f=["=PROCV()"],
        datas_aprovacao={"DOC-POP-1": datetime(2026, 9, 1)},
        linhas_rodape=[linha_do_rodape],
    )
    assert resultado.replay == ["DOC-POP-1"]
    assert not resultado.pode_executar
    assert any("já registrado no rodapé" in b for b in resultado.bloqueios)


# R9 decidido com o negócio em 2026-10-02: versão é comparada como NÚMERO.
def test_e_em_texto_e_f_em_numero_da_mesma_versao_nao_versiona():
    resultado = avaliar_matriz(
        [("DOC-POP-1", "8", 8.0)],
        formulas_f=["=PROCV()"],
        datas_aprovacao={"DOC-POP-1": datetime(2026, 9, 1)},
    )
    assert resultado.alteracoes == {}
    assert resultado.linhas_quitacao == {}


def test_replay_nao_casa_codigo_que_e_sufixo_de_outro():
    alteracoes = {"DOC-POP-1": {"v_antiga": "1.0", "v_nova": "2.0"}}
    rodape = ["Versionado treinamento DOC-CRG-1 de v1.0 para v2.0"]
    assert codigos_com_replay_no_rodape(alteracoes, rodape) == []


# --- Decisões de negócio de 2026-10-02 (R9, E com fórmula, reprocessamento) ---


@pytest.mark.parametrize(
    ("e", "f"),
    [("8", 8.0), ("8.0", 8), (8, "8.0"), ("6.0", 6)],
    ids=["texto 8 x 8.0", "texto 8.0 x 8", "numero 8 x texto 8.0", "Maio/2026 6.0 x 6"],
)
def test_r9_mesma_versao_em_tipos_diferentes_nao_versiona(e, f):
    from treinamentos_its.dominio import avaliar_matriz

    versoes, alteracoes = detectar_alteracoes([("DOC-POP-1", e, f)])
    assert alteracoes == {}
    assert "DOC-POP-1" in versoes
    resultado = avaliar_matriz([("DOC-POP-1", e, f)], formulas_f=["=PROCV()"])
    assert resultado.linhas_quitacao == {}
    assert resultado.pode_executar


def test_r9_versao_maior_continua_versionando():
    _versoes, alteracoes = detectar_alteracoes([("DOC-POP-1", "8.0", 9)])
    assert alteracoes == {"DOC-POP-1": {"v_antiga": "8.0", "v_nova": "9.0"}}


@pytest.mark.parametrize(
    ("e", "f", "esperado"),
    [
        ("1.1", 2.0, "L5 DOC-POP-1 (E='1.1')"),
        (1.0, "1.10", "L5 DOC-POP-1 (F='1.10')"),
        (1.0, 2.5, "L5 DOC-POP-1 (F='2.5')"),
    ],
)
def test_versao_fracionada_trava_o_ciclo(e, f, esperado):
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz([("DOC-POP-1", e, f)], formulas_f=["=PROCV()"])
    assert resultado.fora_do_padrao == [esperado]
    assert not resultado.pode_executar
    assert any("fora do padrão N.0" in b for b in resultado.bloqueios)


def test_versao_inteira_ou_n_ponto_0_nao_trava():
    from treinamentos_its.dominio import versoes_fora_do_padrao

    linhas = [("A-B-1", "8.0", 9), ("A-B-2", 3, "3.0"), ("A-B-3", "2", "VERIFICAR")]
    assert versoes_fora_do_padrao(linhas) == []


def test_e_com_formula_trava_o_ciclo_e_diz_o_que_fazer():
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz(
        [("DOC-POP-1", "12.0", "13.0"), ("DOC-POP-2", "4.0", "4.0"), (None, None, None)],
        formulas_f=["=PROCV()", "=PROCV()", None],
        formulas_e=["=VLOOKUP(C5,X,4,0)", "4.0", "=1+1"],
    )
    assert resultado.e_com_formula == ["L5 DOC-POP-1"]
    assert not resultado.pode_executar
    bloqueio = next(b for b in resultado.bloqueios if b.startswith("E com fórmula"))
    assert "VALOR da versão vigente" in bloqueio
    assert "L5 DOC-POP-1" in bloqueio


def test_e_com_formula_sem_mudanca_de_versao_tambem_trava():
    # E acompanhando F é justamente o caso cego: E == F não prova nada.
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz(
        [("DOC-POP-1", "13.0", "13.0")],
        formulas_f=["=PROCV()"],
        formulas_e=["=VLOOKUP(C5,X,4,0)"],
    )
    assert resultado.alteracoes == {}
    assert resultado.e_com_formula == ["L5 DOC-POP-1"]
    assert not resultado.pode_executar


def test_sem_formulas_e_nao_avalia_trava_da_coluna_e():
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz([("DOC-POP-1", "1.0", "1.0")], formulas_f=["=PROCV()"])
    assert resultado.e_com_formula == []


@pytest.mark.parametrize(
    ("linhas", "esperado"),
    [
        (["Referente Agosto 2026", "Não houveram alterações nos POPs"], True),
        (["  referente   agosto 2026 "], True),
        (["Referente Julho 2026"], False),
        (["Versionado treinamento X Referente Agosto 2026"], False),
        ([None, 42], False),
        (None, False),
    ],
)
def test_rodape_do_ciclo_ja_escrito(linhas, esperado):
    from treinamentos_its.dominio import rodape_do_ciclo_ja_escrito

    assert rodape_do_ciclo_ja_escrito(linhas, "Referente Agosto 2026") is esperado


def test_avaliacao_guarda_as_linhas_do_rodape_normalizadas():
    from treinamentos_its.dominio import avaliar_matriz

    resultado = avaliar_matriz(
        [("DOC-POP-1", "1.0", "1.0")],
        formulas_f=["=PROCV()"],
        linhas_rodape=["  Referente   Agosto 2026 ", None, "", 7],
    )
    assert resultado.linhas_rodape == ["Referente Agosto 2026"]
