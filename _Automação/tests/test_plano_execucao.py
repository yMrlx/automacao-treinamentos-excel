from pathlib import Path

import pytest

from treinamentos_its.ambiente_teste import COMANDOS_COM_TESTE, modo_teste_pedido
from treinamentos_its.cli import construir_parser
from treinamentos_its.plano_execucao import (
    ACAO_CONCLUIR,
    ACAO_CONFERIR,
    ACAO_RESETAR_TESTE,
    ACAO_SYNC,
    ACOES_COM_TESTE,
    CODIGO_SAIDA_RESTAURACAO_INCOMPLETA as CODIGO_SAIDA_PAINEL,
    ESTADO_ATUAL,
    ESTADO_ERRO,
    ESTADO_FEITA,
    ESTADO_NAO_EXECUTADA,
    ESTADO_PENDENTE,
    ETAPAS_CONCLUIR,
    ETAPAS_CONFERIR,
    ETAPAS_RESETAR,
    ETAPAS_SYNC,
    MARCADORES_DE_DESFECHO,
    MODO_PRODUCAO,
    MODO_TESTE,
    NIVEL_ALERTA,
    NIVEL_ERRO,
    NIVEL_ERRO_GRAVE,
    NIVEL_NEUTRO,
    NIVEL_OK,
    TEXTO_NADA_ALTERADO,
    TEXTO_OK,
    TEXTO_RESTAURACAO_FALHOU,
    TEXTO_STATUS_DESCONHECIDOS,
    argumentos_do_comando,
    arquivo_da_linha,
    desfecho_da_execucao,
    estados_das_etapas,
    etapas_da_acao,
    indice_da_etapa,
    motivo_para_nao_rodar,
    plano_de_arquivos,
    resumo_do_plano,
    resumo_humano,
)


MESES = {
    "atual": "Setembro",
    "anterior": "Agosto",
    "ano_atual": 2026,
    "ano_ant_relativo": 2026,
}


@pytest.mark.parametrize("acao", [ACAO_SYNC, ACAO_CONFERIR])
def test_producao_exige_export_de_vigentes(acao):
    assert motivo_para_nao_rodar(acao, MODO_PRODUCAO, vigentes=None)
    assert motivo_para_nao_rodar(acao, MODO_PRODUCAO, vigentes="x.xlsx") is None


def test_argumentos_da_gui_sao_aceitos_pela_cli():
    parser = construir_parser()
    casos = [
        argumentos_do_comando(ACAO_SYNC, MODO_TESTE, mes=8, ano=2026),
        argumentos_do_comando(
            ACAO_CONFERIR, MODO_TESTE, mes=8, ano=2026, vigentes="x.xlsx"
        ),
        argumentos_do_comando(
            ACAO_CONCLUIR,
            MODO_TESTE,
            nome="Carla",
            codigo="DOC-POP-1",
            data="25/09/2026",
        ),
    ]
    for argumentos in casos:
        assert parser.parse_args(argumentos).comando == argumentos[0]


def test_acoes_com_teste_batem_com_as_da_cli():
    assert ACOES_COM_TESTE == COMANDOS_COM_TESTE


@pytest.mark.parametrize("acao", ACOES_COM_TESTE)
def test_producao_nunca_manda_a_flag_de_teste(acao):
    argumentos = argumentos_do_comando(
        acao,
        MODO_PRODUCAO,
        nome="Carla",
        codigo="DOC-POP-1",
        data="25/09/2026",
    )
    assert "--teste" not in argumentos
    assert modo_teste_pedido(["main.py", *argumentos]) is False


@pytest.mark.parametrize("vazio", [None, "", "   "])
def test_caminho_vazio_conta_como_sem_vigentes(vazio):
    assert motivo_para_nao_rodar(ACAO_SYNC, MODO_PRODUCAO, vazio)


def test_modo_desconhecido_e_tratado_como_producao():
    assert motivo_para_nao_rodar(ACAO_SYNC, "modo-invalido", None)


def test_codigo_3_do_painel_bate_com_o_backup():
    from treinamentos_its.backup import CODIGO_SAIDA_RESTAURACAO_INCOMPLETA

    assert CODIGO_SAIDA_PAINEL == CODIGO_SAIDA_RESTAURACAO_INCOMPLETA


def test_plano_do_sync_escreve_so_planos(monkeypatch, tmp_path):
    monkeypatch.setattr("treinamentos_its.plano_execucao.RAIZ_PROJETO", tmp_path)
    le, escreve = plano_de_arquivos(ACAO_SYNC, MODO_PRODUCAO, MESES)
    texto = "\n".join(le + escreve)
    assert "Planos e Macs" in texto
    assert "Procedimentos" not in texto
    assert "_pre-ciclo" not in texto


def test_plano_do_concluir_descreve_ok_e_planejado_vazio():
    texto = "\n".join(
        resumo_humano(
            ACAO_CONCLUIR,
            None,
            modo=MODO_PRODUCAO,
            nome="Carla",
            codigo="DOC-POP-1",
            data="25/09/2026",
        )
    )
    assert "gravo OK" in texto
    assert "esvazio o Planejado" in texto
    assert "Procedimentos" not in texto


def test_plano_do_sync_explica_fonte_unica():
    texto = "\n".join(resumo_humano(ACAO_SYNC, MESES, modo=MODO_TESTE))
    assert "diretamente" in texto
    assert "registro na Matriz a versão nova" in texto
    assert "Procedimentos" not in texto


def test_previa_do_conferir_cita_fallback_de_obsoletos_na_raiz():
    le, _escreve = plano_de_arquivos(ACAO_CONFERIR, MODO_TESTE, MESES)
    detalhes = "\n".join(le)
    resumo = "\n".join(
        resumo_humano(ACAO_CONFERIR, MESES, modo=MODO_TESTE)
    )

    assert "POPs do Sistema" in detalhes and "raiz" in detalhes
    assert "POPs do Sistema" in resumo and "raiz de 'Planos e Macs'" in resumo


@pytest.mark.parametrize(
    "etapas,arquivos",
    [
        (ETAPAS_SYNC, ["backup.py", "vigentes.py", "sync.py", "planos_macs.py"]),
        (ETAPAS_CONFERIR, ["vigentes.py", "sync.py", "conferir.py"]),
        (ETAPAS_CONCLUIR, ["backup.py", "conclusao.py"]),
        (ETAPAS_RESETAR, ["ambiente_teste.py"]),
    ],
)
def test_todo_marcador_de_etapa_existe_no_backend(etapas, arquivos):
    raiz = Path(__file__).resolve().parents[1] / "treinamentos_its"
    codigo = "\n".join((raiz / arquivo).read_text(encoding="utf-8") for arquivo in arquivos)
    for marcador, _rotulo in etapas:
        assert marcador in codigo


def test_marcadores_de_desfecho_existem_no_backend():
    raiz = Path(__file__).resolve().parents[1] / "treinamentos_its"
    codigo = "\n".join(
        (raiz / arquivo).read_text(encoding="utf-8")
        for arquivo in ("cli.py", "conclusao.py", "planos_macs.py", "excel_app.py")
    )
    for marcador in MARCADORES_DE_DESFECHO:
        assert marcador in codigo


@pytest.mark.parametrize(
    "acao,com_vigentes,linha,esperado",
    [
        (ACAO_SYNC, True, "INFO Cópia de restauração de Planos e Macs: C:/x", 0),
        (ACAO_SYNC, True, "INFO IMPORTANDO POPs VIGENTES", 1),
        (ACAO_SYNC, True, "INFO Analisando a Matriz em: C:/x.xlsx", 2),
        (ACAO_SYNC, True, "INFO Matriz avaliada: 2 versionamento(s)", 3),
        (ACAO_SYNC, True, "INFO Versões novas registradas na Matriz (coluna E)", 4),
        (ACAO_SYNC, True, "INFO Planos e Macs conferido após reabrir", 5),
        (ACAO_CONFERIR, True, "INFO IMPORTANDO POPs VIGENTES", 0),
        (ACAO_CONFERIR, True, "INFO Analisando a Matriz em: C:/x.xlsx", 1),
        (ACAO_CONFERIR, True, "INFO Conferência concluída", 2),
        (ACAO_CONCLUIR, False, "INFO Cópia de restauração de Planos e Macs: C:/x", 0),
        (ACAO_CONCLUIR, False, "INFO Conclusão confirmada após reabrir", 1),
        (ACAO_RESETAR_TESTE, False, "INFO Ambiente de teste esvaziado", 0),
        (ACAO_RESETAR_TESTE, False, "INFO sandbox criado com sucesso", 1),
    ],
)
def test_cada_linha_de_log_acende_a_etapa_certa(
    acao, com_vigentes, linha, esperado
):
    assert indice_da_etapa(acao, linha, com_vigentes) == esperado


@pytest.mark.parametrize(
    "acao,linha,esperado",
    [
        (ACAO_SYNC, "INFO Analisando a Matriz em: C:/x.xlsx", 1),
        (ACAO_CONFERIR, "INFO Analisando a Matriz em: C:/x.xlsx", 0),
    ],
)
def test_sem_vigentes_remove_a_etapa_da_importacao(acao, linha, esperado):
    assert indice_da_etapa(acao, linha, com_vigentes=False) == esperado
    assert not any("vigentes" in etapa.lower() for etapa in etapas_da_acao(acao, False))


class TestEstadosDasEtapas:
    def test_rodando_marca_atual_feitas_e_pendentes(self):
        assert estados_das_etapas(4, {0, 1}) == [
            ESTADO_FEITA,
            ESTADO_ATUAL,
            ESTADO_PENDENTE,
            ESTADO_PENDENTE,
        ]

    def test_rodando_etapa_pulada_nao_vira_feita(self):
        assert estados_das_etapas(3, {0, 2}) == [
            ESTADO_FEITA,
            ESTADO_NAO_EXECUTADA,
            ESTADO_ATUAL,
        ]

    def test_nada_aceso_ainda_fica_tudo_pendente(self):
        assert estados_das_etapas(2, set()) == [ESTADO_PENDENTE, ESTADO_PENDENTE]

    def test_sucesso_so_marca_feita_quem_acendeu(self):
        assert estados_das_etapas(4, {0, 2}, terminou=True, sucesso=True) == [
            ESTADO_FEITA,
            ESTADO_NAO_EXECUTADA,
            ESTADO_FEITA,
            ESTADO_NAO_EXECUTADA,
        ]

    def test_sucesso_sem_nada_aceso_nao_tem_check(self):
        assert estados_das_etapas(4, set(), terminou=True, sucesso=True) == [
            ESTADO_NAO_EXECUTADA
        ] * 4

    def test_sucesso_com_tudo_aceso_marca_tudo(self):
        assert estados_das_etapas(3, {0, 1, 2}, terminou=True) == [
            ESTADO_FEITA
        ] * 3

    def test_erro_marca_onde_parou(self):
        assert estados_das_etapas(4, {0, 1}, terminou=True, sucesso=False) == [
            ESTADO_FEITA,
            ESTADO_ERRO,
            ESTADO_NAO_EXECUTADA,
            ESTADO_NAO_EXECUTADA,
        ]

    def test_erro_sem_nada_aceso_marca_a_primeira(self):
        assert estados_das_etapas(3, set(), terminou=True, sucesso=False) == [
            ESTADO_ERRO,
            ESTADO_NAO_EXECUTADA,
            ESTADO_NAO_EXECUTADA,
        ]

    def test_indices_fora_da_faixa_sao_ignorados(self):
        assert estados_das_etapas(2, {5, -1}, terminou=True) == [
            ESTADO_NAO_EXECUTADA
        ] * 2

    def test_zero_etapas(self):
        assert estados_das_etapas(0, {0}, terminou=True) == []


def test_desfechos_honestos():
    assert desfecho_da_execucao(ACAO_SYNC, 0, ["Nada foi alterado"] ) == (
        TEXTO_NADA_ALTERADO,
        NIVEL_NEUTRO,
    )
    assert desfecho_da_execucao(ACAO_SYNC, 0, ["ok"]) == (TEXTO_OK, NIVEL_OK)
    assert desfecho_da_execucao(ACAO_SYNC, 1, []) [1] == NIVEL_ERRO
    assert desfecho_da_execucao(ACAO_SYNC, 3, []) == (
        TEXTO_RESTAURACAO_FALHOU,
        NIVEL_ERRO_GRAVE,
    )


@pytest.fixture
def avisos_status_desconhecidos(caplog):
    from treinamentos_its import planos_macs

    # Usa o log realmente emitido pelo backend, não uma cópia da mensagem.
    planos_macs._avisar_status_desconhecidos(
        Path("Planos e Macs Setembro 2026.xlsx"),
        ["Operador!G7 | Carla | DOC-POP-1 | INDEFINIDO"],
    )
    return caplog.text.splitlines()


def test_sync_com_aviso_final_do_backend_termina_em_alerta(avisos_status_desconhecidos):
    # Uma linha posterior e um iterable de passagem única não apagam o aviso.
    linhas = iter([*avisos_status_desconhecidos, "INFO Fim do processo."])
    assert desfecho_da_execucao(ACAO_SYNC, 0, linhas) == (
        TEXTO_STATUS_DESCONHECIDOS, NIVEL_ALERTA,
    )


def test_sync_sem_status_desconhecido_continua_verde(caplog):
    from treinamentos_its import planos_macs

    planos_macs._avisar_status_desconhecidos(Path("Setembro.xlsx"), [])
    assert caplog.text == ""
    assert desfecho_da_execucao(ACAO_SYNC, 0, caplog.text.splitlines()) == (
        TEXTO_OK, NIVEL_OK,
    )


@pytest.mark.parametrize("acao", [ACAO_CONFERIR, ACAO_CONCLUIR, ACAO_RESETAR_TESTE])
def test_alerta_de_status_e_exclusivo_do_sync(acao, avisos_status_desconhecidos):
    assert desfecho_da_execucao(acao, 0, avisos_status_desconhecidos) == (
        TEXTO_OK, NIVEL_OK,
    )


@pytest.mark.parametrize("codigo,nivel", [(1, NIVEL_ERRO), (3, NIVEL_ERRO_GRAVE)])
def test_aviso_de_status_nao_encobre_erro(codigo, nivel, avisos_status_desconhecidos):
    assert desfecho_da_execucao(ACAO_SYNC, codigo, avisos_status_desconhecidos)[1] == nivel


def test_aviso_de_status_nao_encobre_restauracao_incompleta(avisos_status_desconhecidos):
    linhas = ["ERROR RESTAURAÇÃO INCOMPLETA", *avisos_status_desconhecidos]
    assert desfecho_da_execucao(ACAO_SYNC, 0, linhas) == (
        TEXTO_RESTAURACAO_FALHOU, NIVEL_ERRO_GRAVE,
    )


def test_nada_alterado_mantem_prioridade_sobre_aviso_de_status(avisos_status_desconhecidos):
    linhas = [*avisos_status_desconhecidos, "INFO Nada foi alterado."]
    assert desfecho_da_execucao(ACAO_SYNC, 0, linhas) == (
        TEXTO_NADA_ALTERADO, NIVEL_NEUTRO,
    )


@pytest.mark.parametrize("linha", ["WARNING aviso qualquer", "INFO Atualizado com sucesso."])
def test_sync_sem_o_marcador_especifico_nao_vira_alerta(linha):
    assert desfecho_da_execucao(ACAO_SYNC, 0, [linha]) == (TEXTO_OK, NIVEL_OK)


def test_arquivo_da_linha_pega_o_primeiro_xlsx():
    assert arquivo_da_linha(
        r"INFO Criado: C:\x\Planos e Macs Setembro 2026.xlsx (cópia de Agosto.xlsx)"
    ) == "Planos e Macs Setembro 2026.xlsx"


# --- texto_da_pergunta: caixa de pergunta limpa (05/10/2026) ---

LOG_DA_GUARDA = [
    "• No fim, registro na Matriz a versão nova como a versão atual.",
    "$ python main.py sync --mes 8 --ano 2026 --vigentes C:/x.xlsx --teste",
    "05/10/2026 14:26:01  INFO      Mês de referência informado: Agosto/2026",
    "05/10/2026 14:26:01  INFO      =======================================================",
    "05/10/2026 14:26:01  INFO      Automação de Treinamentos - Time de TI",
    "05/10/2026 14:26:06  WARNING   O ciclo de Setembro/2026 parece JÁ ter sido processado:",
    "05/10/2026 14:26:06  WARNING     - a planilha 'Planos e Macs Setembro 2026.xlsx' de destino já existe",
    "05/10/2026 14:26:06  WARNING     - o rodapé 'Referente Agosto 2026' já está escrito em: Analista",
    "Rodar mesmo assim? (s/n): ",
]


def test_pergunta_da_guarda_mostra_so_motivos_pergunta_e_botoes():
    from treinamentos_its.plano_execucao import EXPLICACAO_REPROCESSAR, texto_da_pergunta

    texto = texto_da_pergunta(LOG_DA_GUARDA)

    assert texto == (
        "O ciclo de Setembro/2026 parece JÁ ter sido processado:\n"
        "  - a planilha 'Planos e Macs Setembro 2026.xlsx' de destino já existe\n"
        "  - o rodapé 'Referente Agosto 2026' já está escrito em: Analista\n\n"
        "Rodar mesmo assim? (s/n):\n\n" + EXPLICACAO_REPROCESSAR
    )
    assert "14:26" not in texto and "INFO" not in texto and "$ python" not in texto


def test_pergunta_sem_warning_antes_mostra_so_a_pergunta():
    from treinamentos_its.plano_execucao import texto_da_pergunta

    texto = texto_da_pergunta(
        ["05/10/2026 14:26:01  INFO      algo", "Continuar? (s/n): "]
    )
    assert texto == "Continuar? (s/n):"


def test_sem_pergunta_reconhecida_cai_nas_ultimas_linhas():
    from treinamentos_its.plano_execucao import texto_da_pergunta

    linhas = [f"linha {i}" for i in range(10)]
    assert texto_da_pergunta(linhas) == "\n".join(linhas[-6:])


# --- Entradas da operadora (QA de 05/10/2026) ---


@pytest.mark.parametrize(
    ("mes", "ano", "esperado"),
    [
        (8, 2026, None),
        (8, 9999, "O ano tem que estar entre 2020 e 2100."),
        (8, 1900, "O ano tem que estar entre 2020 e 2100."),
        (13, 2026, "O mês tem que estar entre 1 e 12."),
        (0, 2026, "O mês tem que estar entre 1 e 12."),
        (None, 2026, "Preencha o mês (1 a 12) e o ano com números."),
        (8, None, "Preencha o mês (1 a 12) e o ano com números."),
    ],
)
def test_problema_no_mes_escolhido(mes, ano, esperado):
    from treinamentos_its.plano_execucao import problema_no_mes_escolhido

    assert problema_no_mes_escolhido(mes, ano) == esperado


def _meses_de_setembro():
    # POPs de Setembro: cria o Planos e Macs de Setembro a partir do de Agosto.
    from treinamentos_its.sync import obter_meses

    return obter_meses(9, 2026)


def test_export_do_mes_certo_nao_avisa():
    from treinamentos_its.plano_execucao import avisos_de_mes_dos_exports

    assert avisos_de_mes_dos_exports(
        _meses_de_setembro(), {"vigentes": (9, 2026), "obsoletos": (9, None)}
    ) == []


def test_export_de_outro_mes_ou_ano_avisa_cada_um():
    from treinamentos_its.plano_execucao import avisos_de_mes_dos_exports

    avisos = avisos_de_mes_dos_exports(
        _meses_de_setembro(), {"vigentes": (8, 2026), "obsoletos": (9, 2025)}
    )
    assert len(avisos) == 2
    assert avisos[0].startswith("O export de POPs vigentes parece ser de Agosto/2026")
    assert (
        "mas o mês escolhido é Setembro/2026 (POPs de Setembro criam o Planos e Macs "
        "de Setembro a partir do de Agosto)" in avisos[0]
    )
    assert "obsoletos parece ser de Setembro/2025" in avisos[1]


def test_nome_sem_mes_ou_sem_ciclo_nao_avisa():
    from treinamentos_its.plano_execucao import avisos_de_mes_dos_exports

    assert avisos_de_mes_dos_exports(_meses_de_setembro(), {"vigentes": (None, None)}) == []
    assert avisos_de_mes_dos_exports(None, {"vigentes": (9, 2026)}) == []


# --- Fim do ciclo: abrir o arquivo criado (pedido da operadora, 06/10/2026) ---


def test_arquivo_criado_pelo_ciclo_e_o_do_mes_dos_pops():
    from treinamentos_its.plano_execucao import arquivo_criado_pelo_ciclo
    from treinamentos_its.sync import obter_meses

    caminho = arquivo_criado_pelo_ciclo(MODO_PRODUCAO, obter_meses(9, 2026))
    assert caminho.name == "Planos e Macs Setembro 2026.xlsx"
    assert caminho.parent.name == "Planos e Macs"
    assert arquivo_criado_pelo_ciclo(MODO_PRODUCAO, None) is None


GRAVOU = "INFO      Planos e Macs atualizado com sucesso."
NADA = "INFO      Nada novo para registrar: ... Nada foi alterado no Planos e Macs de Setembro."
EXCEL_VIVO = "ERROR     NÃO consegui encerrar o Excel invisível (processo 123)."


@pytest.mark.parametrize(
    ("acao", "nivel", "log", "oferece"),
    [
        (ACAO_SYNC, "ok", [GRAVOU], True),
        (ACAO_SYNC, "alerta", [GRAVOU, "WARNING   3 status desconhecido(s)"], True),
        (ACAO_SYNC, "alerta", [GRAVOU, EXCEL_VIVO], True),  # gravou; o alerta continua
        (ACAO_SYNC, "neutro", [NADA], False),
        (ACAO_SYNC, "erro", [GRAVOU], False),
        (ACAO_SYNC, "erro_grave", [GRAVOU], False),
        (ACAO_CONFERIR, "ok", [GRAVOU], False),
        (ACAO_CONCLUIR, "ok", [GRAVOU], False),
        # Sem prova de gravação no log, nível OK sozinho não basta.
        (ACAO_SYNC, "ok", ["INFO      Terminou."], False),
    ],
)
def test_so_oferece_abrir_depois_do_ciclo_que_gravou(acao, nivel, log, oferece):
    from treinamentos_its.plano_execucao import oferecer_abrir_o_arquivo

    assert oferecer_abrir_o_arquivo(acao, nivel, log) is oferece


@pytest.mark.parametrize(
    "log",
    [
        # Reprocesso recusado ("n") com o Excel da guarda que não fechou.
        ["INFO      Nada foi alterado.", EXCEL_VIVO],
        # Reprocesso sem versão nova, desfeito pela transação, com Excel vivo.
        [NADA, EXCEL_VIVO],
    ],
)
def test_alerta_de_excel_vivo_sem_gravacao_nao_oferece_abrir(log):
    """Revisão do Codex (06/10/2026): o desfecho vira ALERTA (o Excel vivo vem
    antes do "Nada foi alterado"), mas nada foi gravado - não pode oferecer
    abrir o arquivo como se o ciclo tivesse produzido algo."""
    from treinamentos_its.plano_execucao import (
        desfecho_da_execucao,
        oferecer_abrir_o_arquivo,
    )

    _texto, nivel = desfecho_da_execucao(ACAO_SYNC, 0, log)
    assert nivel == "alerta"
    assert oferecer_abrir_o_arquivo(ACAO_SYNC, nivel, log) is False
