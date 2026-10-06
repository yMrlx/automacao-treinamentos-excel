from treinamentos_its.limpeza_rodape import mesclar_blocos_duplicados


def test_sem_duplicata_nao_mexe_em_nada():
    linhas = [
        "cabecalho", "dado1", None,
        "Referente Junho 2026", "Versionado treinamento COD1 de v1.0 para v2.0",
        None,
        "Referente Julho 2026", "Não houveram alterações nos POPs",
    ]

    assert mesclar_blocos_duplicados(linhas) == linhas


def test_bloco_unico_nao_mexe():
    linhas = [
        None, "cabecalho", "dado1", None,
        "Referente Julho 2026",
        "Versionado do treinamento COD1 de v1.0 para v2.0",
        "FULANO - COD2 ATRASADO",
    ]

    assert mesclar_blocos_duplicados(linhas) == linhas


def test_dois_blocos_mesmo_versionado_atrasado_diferente_mescla():
    linhas = [
        "dado", None,
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "FULANO - COD1 ATRASADO",
        None,
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "CICRANO - COD1 ATRASADO",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "dado", None,
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "CICRANO - COD1 ATRASADO",  # só o bloco mais recente (último)
    ]


def test_dois_blocos_versionado_diferente_une_sem_duplicar_codigo():
    linhas = [
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        None,
        "Referente Junho 2026",
        "Versionado treinamento COD2 de v3.0 para v4.0",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "Versionado treinamento COD2 de v3.0 para v4.0",
    ]


def test_tres_blocos_empilhados_mescla_os_tres():
    linhas = [
        "Referente Julho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "FULANO - COD1 ATRASADO",
        None,
        "Referente Julho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "Versionado treinamento COD2 de v5.0 para v6.0",
        "FULANO - COD1 ATRASADO",
        "CICRANO - COD2 ATRASADO",
        None,
        "Referente Julho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "Versionado treinamento COD2 de v5.0 para v6.0",
        "BELTRANO - COD2 ATRASADO",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "Referente Julho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        "Versionado treinamento COD2 de v5.0 para v6.0",
        "BELTRANO - COD2 ATRASADO",  # só a lista de ATRASADO do 3º (último) bloco
    ]


def test_sem_versionado_sobrando_mantem_nao_houveram_alteracoes():
    linhas = [
        "Referente Julho 2026", "Não houveram alterações nos POPs", "FULANO - COD1 ATRASADO",
        None,
        "Referente Julho 2026", "Não houveram alterações nos POPs", "CICRANO - COD1 ATRASADO",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
        "CICRANO - COD1 ATRASADO",
    ]


def test_linhas_verificar_mesclam_igual_versionado_sem_duplicar_codigo():
    linhas = [
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
        "COD1 — POP obsoletado ou renomeado, para um humano verificar",
        None,
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
        "COD1 — POP obsoletado ou renomeado, para um humano verificar",
        "COD2 — POP obsoletado ou renomeado, para um humano verificar",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
        "COD1 — POP obsoletado ou renomeado, para um humano verificar",
        "COD2 — POP obsoletado ou renomeado, para um humano verificar",
    ]


def test_titulos_diferentes_nao_se_misturam():
    # Duplicata só em "Referente Junho 2026"; "Referente Julho 2026" (único) fica intocado.
    linhas = [
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        None,
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        None,
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
    ]

    resultado = mesclar_blocos_duplicados(linhas)

    assert resultado == [
        "Referente Junho 2026",
        "Versionado treinamento COD1 de v1.0 para v2.0",
        None,
        "Referente Julho 2026",
        "Não houveram alterações nos POPs",
    ]


def test_sem_titulo_nenhum_devolve_igual():
    linhas = [None, "cabecalho", "dado1", "dado2"]
    assert mesclar_blocos_duplicados(linhas) == linhas
