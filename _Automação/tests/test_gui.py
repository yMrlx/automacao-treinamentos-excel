"""Testes da casca tkinter: o que trava, o que roda, o que ela diz no fim.

Cria a janela de verdade (escondida), mas NUNCA roda o programa: o `_rodar` é
trocado por um gravador e o `subprocess.Popen` explode se alguém chamar. As
caixas de diálogo também são dublês - nada fica esperando clique.

UMA janela pro arquivo inteiro, com o estado zerado antes de cada teste. Criar
um `Tk()` por teste falhava de vez em quando nesta máquina ("Can't find a usable
tk.tcl", lá pela 9ª janela da suíte) e o teste virava "pulado" em silêncio.

Se esta máquina não tiver Tk (ambiente sem tela), os testes são pulados.
"""

import tkinter as tk
from datetime import datetime

import pytest

from treinamentos_its import gui
from treinamentos_its.plano_execucao import (
    ACAO_CONCLUIR,
    ACAO_CONFERIR,
    ACAO_SYNC,
    MODO_PRODUCAO,
    MODO_TESTE,
    MARCADOR_STATUS_DESCONHECIDOS,
    TEXTO_NADA_ALTERADO,
    TEXTO_OK,
    TEXTO_RESTAURACAO_FALHOU,
    TEXTO_STATUS_DESCONHECIDOS,
    etapas_da_acao,
    resumo_humano,
)

DADOS_CONCLUSAO = {"nome": "Fulana de Tal", "codigo": "DOC-POP-1", "data": "01/09/2026"}


def _log(nivel, mensagem):
    return f"25/09/2026 10:00:00  {nivel:<8}  {mensagem}"


def _nunca_rodar(*_args, **_kwargs):
    raise AssertionError("teste de GUI tentou abrir um subprocesso de verdade")


@pytest.fixture(scope="module")
def _janela():
    ultimo_erro = None
    for _tentativa in range(3):
        try:
            janela = gui.Painel()
            break
        except tk.TclError as erro:  # sem tela / sem Tk / tk.tcl intermitente
            ultimo_erro = erro
    else:
        pytest.skip(f"Tk indisponível nesta máquina: {ultimo_erro}")
    janela.withdraw()
    yield janela
    janela.fechando = True
    janela.destroy()


def _zerar(janela) -> None:
    """Volta a janela ao estado de recém-aberta."""
    janela.processo = None
    janela.acao_em_curso = None
    janela.linha_parcial = ""
    janela.cauda = ""
    janela.modo.set(MODO_TESTE)
    janela.usar_mes_manual.set(False)
    janela.mes_manual.set(str(datetime.now().month))
    janela.ano_manual.set(str(datetime.now().year))
    janela.arquivo_vigentes.set("")
    janela.arquivo_obsoletos.set("")
    janela._travar_botoes(False)
    janela.saida.delete("1.0", "end")
    janela._pintar_status("Parado.")
    janela._atualizar_modo()


@pytest.fixture
def painel(_janela, monkeypatch):
    monkeypatch.setattr(gui.subprocess, "Popen", _nunca_rodar)
    _zerar(_janela)
    return _janela


@pytest.fixture
def caixas(monkeypatch):
    """Grava toda caixa de diálogo; as de pergunta respondem SIM."""
    chamadas: list[tuple[str, str, str]] = []

    def gravador(tipo, resposta):
        def caixa(titulo="", mensagem="", **_kwargs):
            chamadas.append((tipo, titulo, mensagem))
            return resposta
        return caixa

    for tipo in ("showwarning", "showerror", "showinfo"):
        monkeypatch.setattr(gui.messagebox, tipo, gravador(tipo, "ok"))
    for tipo in ("askokcancel", "askyesno"):
        monkeypatch.setattr(gui.messagebox, tipo, gravador(tipo, True))
    return chamadas


@pytest.fixture
def execucoes(painel, monkeypatch):
    """O que teria ido pro subprocesso, em vez de ir."""
    chamadas: list[tuple[list[str], str]] = []
    monkeypatch.setattr(
        painel, "_rodar", lambda argumentos, acao: chamadas.append((argumentos, acao))
    )
    return chamadas


def _saida(painel) -> str:
    return painel.saida.get("1.0", "end")


def _marcas(painel) -> list[str]:
    return [rotulo.cget("text")[0] for rotulo in painel.rotulos_etapas]


class TestTravaDuranteExecucao:
    def test_modo_mes_e_arquivos_ficam_travados_enquanto_roda(self, painel):
        controles = painel.controles_de_configuracao
        # 2 rádios de modo + checkbox + 2 spinbox + (Escolher, Tirar) x 2 arquivos
        assert len(controles) == 9
        assert all(radio in controles for radio in painel.radios_modo)
        assert painel.check_mes in controles
        assert painel.spin_mes in controles and painel.spin_ano in controles

        painel._travar_botoes(True)
        assert all(controle.instate(["disabled"]) for controle in controles)
        assert all(botao.instate(["disabled"]) for botao in painel.botoes.values())

        painel._travar_botoes(False)
        assert not any(controle.instate(["disabled"]) for controle in controles)

    def test_fim_da_execucao_destrava(self, painel):
        painel._iniciar_andamento(ACAO_SYNC, False)
        painel._travar_botoes(True)
        painel._terminou(0)
        assert not any(c.instate(["disabled"]) for c in painel.controles_de_configuracao)

    def test_falha_ao_abrir_o_processo_destrava_e_avisa(self, painel, monkeypatch):
        def falha(*_args, **_kwargs):
            raise OSError("python sumiu")

        monkeypatch.setattr(gui.subprocess, "Popen", falha)
        painel._iniciar_andamento(ACAO_CONFERIR, False)
        painel._rodar([ACAO_CONFERIR, "--teste"], ACAO_CONFERIR)
        assert painel.processo is None and painel.acao_em_curso is None
        assert not any(c.instate(["disabled"]) for c in painel.controles_de_configuracao)
        assert "Não consegui iniciar o programa" in _saida(painel)
        assert "erro" in painel.rotulo_status.cget("text")


class TestProducaoExigeVigentes:
    @pytest.mark.parametrize("acao", [ACAO_SYNC, ACAO_CONFERIR])
    def test_producao_sem_vigentes_nao_roda(self, painel, caixas, execucoes, acao):
        painel.modo.set(MODO_PRODUCAO)
        painel._atualizar_modo()
        painel._executar(acao)
        assert execucoes == []
        avisos = [c for c in caixas if c[0] == "showwarning"]
        assert len(avisos) == 1
        assert "POPs vigentes" in avisos[0][2]
        assert "Matriz antiga do mês anterior" in avisos[0][2]

    def test_teste_sem_vigentes_roda_o_conferir(self, painel, caixas, execucoes):
        painel.modo.set(MODO_TESTE)
        painel._executar(ACAO_CONFERIR)
        assert len(execucoes) == 1
        argumentos, acao = execucoes[0]
        assert acao == ACAO_CONFERIR
        assert "--teste" in argumentos and "--pops-vigentes" not in argumentos
        # sem export, a etapa da cópia descartável nem aparece
        assert painel.textos_etapas == etapas_da_acao(ACAO_CONFERIR, com_vigentes=False)


class TestPreviaNaSaida:
    def test_conferir_escreve_a_propria_previa_antes_de_rodar(self, painel, caixas, execucoes):
        painel.modo.set(MODO_TESTE)
        painel._executar(ACAO_CONFERIR)
        saida = _saida(painel)
        for frase in resumo_humano(ACAO_CONFERIR, painel._meses(), modo=MODO_TESTE):
            assert frase in saida
        assert "Conferir antes de rodar" in saida and "TESTE" in saida
        # o conferir não escreve: nenhuma caixa de confirmação
        assert caixas == []

    def test_previa_do_quadro_diz_que_producao_sem_vigentes_nao_roda(self, painel):
        painel.modo.set(MODO_PRODUCAO)
        painel._atualizar_modo()
        assert "só rodo com o arquivo de POPs vigentes" in painel.rotulo_previa.cget("text")
        painel.modo.set(MODO_TESTE)
        painel._atualizar_modo()
        assert "formato ANTIGO" in painel.rotulo_previa.cget("text")


class TestConcluir:
    def test_concluir_ignora_o_mes_do_ciclo(self, painel, caixas, execucoes, monkeypatch):
        painel.usar_mes_manual.set(True)
        painel.mes_manual.set("8")
        painel.ano_manual.set("2026")
        monkeypatch.setattr(painel, "_perguntar_conclusao", lambda: dict(DADOS_CONCLUSAO))
        painel._executar(ACAO_CONCLUIR)
        assert len(execucoes) == 1
        argumentos, _acao = execucoes[0]
        assert "--mes" not in argumentos and "--ano" not in argumentos
        assert argumentos[:3] == ["concluir", "Fulana de Tal", "DOC-POP-1"]
        assert ["--data", "01/09/2026"] == argumentos[3:5]
        # a confirmação (antes estourava TypeError por causa do `data`) apareceu
        assert [c[0] for c in caixas] == ["askokcancel"]
        assert "Setembro/2026" in caixas[0][2]
        assert "Setembro/2026" in _saida(painel)

    def test_mes_invalido_no_seletor_nao_trava_o_concluir(
        self, painel, caixas, execucoes, monkeypatch
    ):
        painel.usar_mes_manual.set(True)
        painel.mes_manual.set("abc")
        monkeypatch.setattr(painel, "_perguntar_conclusao", lambda: dict(DADOS_CONCLUSAO))
        painel._executar(ACAO_CONCLUIR)
        assert len(execucoes) == 1


class TestDesfecho:
    def test_status_desconhecido_mostra_amarelo_so_com_checks_reais(self, painel):
        painel._iniciar_andamento(ACAO_SYNC, False)
        painel._processar_linha(_log("INFO", "Cópia de restauração de Planos e Macs: C:/x"))
        # Também cobre o aviso como última linha sem quebra do subprocesso.
        painel.linha_parcial = _log("WARNING", f"1 {MARCADOR_STATUS_DESCONHECIDOS} em Setembro.xlsx")
        painel._terminou(0)
        assert painel.rotulo_status.cget("text") == TEXTO_STATUS_DESCONHECIDOS
        assert painel.rotulo_status.cget("bg").upper() == "#FFD54F"
        assert _marcas(painel)[0] == gui.MARCA_FEITA
        assert set(_marcas(painel)[1:]) == {gui.MARCA_NAO_EXECUTADA}
        assert TEXTO_STATUS_DESCONHECIDOS.rstrip(".") in _saida(painel)

    def test_guarda_respondida_n_nao_marca_nada_como_feito(self, painel):
        painel._iniciar_andamento(ACAO_SYNC, False)
        for linha in (
            _log("WARNING", "O ciclo de Setembro/2026 parece JÁ ter sido processado:"),
            "Rodar mesmo assim? (s/n): " + _log("INFO", "Ok, não vou reprocessar. Nada foi alterado."),
        ):
            painel._processar_linha(linha)
        painel._terminou(0)
        assert painel.rotulo_status.cget("text") == TEXTO_NADA_ALTERADO
        assert set(_marcas(painel)) == {gui.MARCA_NAO_EXECUTADA}
        assert "=== Nada foi alterado ===" in _saida(painel)

    def test_restauracao_falhou_e_erro_grave(self, painel):
        painel._iniciar_andamento(ACAO_SYNC, False)
        painel._processar_linha(_log("INFO", "Cópia de restauração de Planos e Macs: C:/x"))
        painel._processar_linha(_log("ERROR", "RESTAURAÇÃO FALHOU para Planos e Macs"))
        painel._processar_linha(_log("ERROR", "RESTAURAÇÃO INCOMPLETA: 1 arquivo(s) ..."))
        painel._terminou(3)
        assert painel.rotulo_status.cget("text") == TEXTO_RESTAURACAO_FALHOU
        assert painel.rotulo_status.cget("bg").upper() == "#B00020"
        assert _marcas(painel)[0] == gui.MARCA_ERRO

    def test_ultima_linha_sem_quebra_ainda_acende_a_etapa(self, painel):
        # A sobra de linha era processada DEPOIS de zerar a ação - e ignorada.
        painel._iniciar_andamento(ACAO_CONCLUIR, False)
        painel._processar_linha(
            _log("INFO", "Cópia de restauração de Planos e Macs durante a conclusão: C:/x")
        )
        painel.linha_parcial = _log(
            "INFO", "Conclusão confirmada após reabrir: Carla + DOC-POP-1"
        )
        painel._terminou(0)
        assert painel.rotulo_status.cget("text") == TEXTO_OK
        assert _marcas(painel) == [gui.MARCA_FEITA, gui.MARCA_FEITA]

    def test_erro_comum(self, painel):
        painel._iniciar_andamento(ACAO_SYNC, False)
        painel._terminou(1)
        assert "código 1" in painel.rotulo_status.cget("text")
        assert _marcas(painel)[0] == gui.MARCA_ERRO
        assert set(_marcas(painel)[1:]) == {gui.MARCA_NAO_EXECUTADA}


# --- Painel da operadora: só PRODUÇÃO (05/10/2026) ---


def test_argumento_do_bat_liga_o_modo_somente_producao():
    assert gui.abre_somente_em_producao(["--somente-producao"]) is True
    assert gui.abre_somente_em_producao([]) is False
    assert gui.abre_somente_em_producao(["--outra-coisa"]) is False


def test_producao_e_turquesa_da_marca_e_teste_nao_e_parecido():
    assert gui.COR_PRODUCAO_FUNDO.upper() == "#00AEAB"
    assert gui.COR_TESTE_FUNDO.upper() not in ("#00AEAB", "#0B6E3A")


def test_painel_da_operadora_so_tem_producao_e_sem_reset():
    from tkinter import ttk

    from treinamentos_its.plano_execucao import ACAO_RESETAR_TESTE

    try:
        janela = gui.Painel(somente_producao=True)
    except tk.TclError as erro:
        pytest.skip(f"Tk indisponível nesta máquina: {erro}")
    try:
        janela.withdraw()
        assert janela.modo.get() == MODO_PRODUCAO
        assert janela.radios_modo == []
        assert ACAO_RESETAR_TESTE not in janela.botoes
        assert set(janela.botoes) == {ACAO_CONFERIR, ACAO_SYNC, ACAO_CONCLUIR}
        titulos = [
            filho.cget("text").strip()
            for filho in janela.winfo_children()
            if isinstance(filho, ttk.LabelFrame)
        ]
        assert titulos[:3] == [
            "1. Os dois arquivos que chegaram por e-mail",
            "2. Mês dos POPs (o Planos e Macs que vai ser criado)",
            "3. O que fazer",
        ]
        assert janela.banner.cget("bg").upper() == "#00AEAB"
    finally:
        janela.fechando = True
        janela.destroy()
