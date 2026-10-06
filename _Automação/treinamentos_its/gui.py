"""Janela única que substitui os cinco ``.bat`` do dia a dia.

Esta camada é **casca**: ela não sabe nada de Matriz, versão ou prazo. Tudo o que
faz é (1) montar a lista de argumentos com ``plano_execucao.argumentos_do_comando``,
(2) rodar ``main.py`` num **subprocesso** e (3) despejar a saída na tela. Regra de
negócio nenhuma mora aqui — nem a de "o que dizer no fim", que é
``plano_execucao.desfecho_da_execucao``.

Por que subprocesso e não ``import cli``: o ``config.PATH_BASE`` é resolvido uma
única vez, no import do ``config``. Num processo só, alternar entre teste e
produção exigiria reimportar meio pacote, e bastaria um import esquecido pra
janela dizer "teste" enquanto escreve nas planilhas reais. Cada execução sai num
processo novo, que nasce com a base certa e passa pelas MESMAS travas do
``main.py`` (``modo_teste_pedido``) e do ``cli`` (``_abortar_se_teste_sem_sandbox``).
A GUI é mais uma camada por fora, não um atalho por dentro.

O trabalho pesado (o Excel abrindo, o ciclo rodando) fica no subprocesso e a
leitura da saída numa thread; a interface só recebe texto por uma fila. Assim a
janela não congela — congelada, a operadora acharia que travou e fecharia no
meio, que é como se deixa planilha pela metade.
"""

from __future__ import annotations

import codecs
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from treinamentos_its.ambiente_teste import PASTA_AUTOMACAO, PASTA_SANDBOX
from treinamentos_its.logger import log
from treinamentos_its.plano_execucao import (
    ACAO_CONCLUIR,
    ACAO_CONFERIR,
    ACAO_RESETAR_TESTE,
    ACAO_SYNC,
    ACOES_COM_MES,
    ACOES_COM_TESTE,
    ESTADO_ATUAL,
    ESTADO_ERRO,
    ESTADO_FEITA,
    ESTADO_NAO_EXECUTADA,
    ESTADO_PENDENTE,
    MODO_PRODUCAO,
    MODO_TESTE,
    NIVEIS_SEM_ERRO,
    NIVEL_ALERTA,
    NIVEL_ERRO,
    NIVEL_ERRO_GRAVE,
    NIVEL_NEUTRO,
    NIVEL_OK,
    acao_escreve_em_arquivo,
    argumentos_do_comando,
    arquivo_criado_pelo_ciclo,
    arquivo_da_linha,
    avisos_de_mes_dos_exports,
    base_do_modo,
    desfecho_da_execucao,
    estados_das_etapas,
    etapas_da_acao,
    indice_da_etapa,
    motivo_para_nao_rodar,
    oferecer_abrir_o_arquivo,
    plano_de_arquivos,
    problema_no_mes_escolhido,
    resumo_humano,
    texto_da_pergunta,
    texto_do_banner,
)
from treinamentos_its.vigentes import descrever_mes_inferido, inferir_mes_ano

# --- cores -------------------------------------------------------------------
#
# Duas famílias de cor, com papéis DIFERENTES, e elas não se misturam:
#
# 1. **Identidade (marca).** `#00AEAB`, o turquesa institucional. Aparece em
#    TRAÇOS FINOS e em TEXTO - régua da faixa de marca, título, barra de
#    progresso. Nunca como bloco grande preenchido. É decoração.
# 2. **Sinal de modo (teste x produção).** No letreiro grande e na moldura da
#    janela inteira. É a defesa contra o pior bug que este projeto teve -
#    alguém achar que está testando e estar escrevendo nas planilhas reais.
#    **Desde 05/10/2026 (pedido do responsável pelo processo) a PRODUÇÃO é o turquesa da marca**
#    (era vermelho) e o TESTE virou cinza-azulado (era verde): turquesa e verde
#    são vizinhos (2,31:1, 30° de matiz) e no painel completo os dois modos
#    ficariam parecidos. O painel da operadora (`--somente-producao`) nem tem
#    TESTE. Texto escuro sobre o turquesa: branco nele dá só 2,75:1.
#
# **Por que a faixa de marca é BRANCA e não turquesa.** O letreiro do modo
# teste é verde (`#0B6E3A`) e o turquesa é vizinho dele: medido, dá razão de
# contraste 2,31:1 com só 30° de diferença de matiz. Uma faixa turquesa cheia
# logo acima do letreiro verde viraria "dois verdes empilhados", e o pior
# resultado possível aqui é alguém não distinguir a marca do aviso de segurança.
# Com a faixa branca a distinção vai pra 6,35:1 e 148° de matiz - some a dúvida.
# A marca cede, o sinal de perigo não. (O oliva anterior separava melhor, 89°;
# o turquesa é a cor certa, então quem se adapta é o desenho.)
COR_MARCA = "#00AEAB"
COR_FAIXA_MARCA = "#FFFFFF"
# Turquesa escurecido pra texto: `#00AEAB` puro sobre branco é claro demais.
# Este dá 5,93:1 sobre branco e 5,20:1 sobre o cinza do tema do ttk.
COR_MARCA_ESCURA = "#00706E"

COR_TESTE_FUNDO = "#455A64"
COR_TEXTO_TESTE = "#FFFFFF"  # 7,2:1 sobre o cinza-azulado
COR_PRODUCAO_FUNDO = COR_MARCA
COR_TEXTO_PRODUCAO = "#002E2D"  # 6,3:1 sobre o turquesa
COR_FUNDO_SAIDA = "#111418"
COR_TEXTO_SAIDA = "#E6E6E6"

# Logo institucional, opcional. PNG porque o tkinter abre nativo (`PhotoImage`),
# sem depender do Pillow - uma dependência a mais só pra enfeite não se paga.
# Se o arquivo não existir ou não der pra ler, o painel abre sem ele: é recurso
# de apresentação, não pode derrubar a ferramenta na mão da operadora.
NOME_ARQUIVO_LOGO = "logo.png"
ALTURA_MAX_LOGO = 48

# O `input()` do `perguntar_sim_nao` escreve o prompt sem quebra de linha. Como a
# saída é lida em pedaços, a pergunta é reconhecida pelo fim do texto acumulado.
PADRAO_PERGUNTA = re.compile(r"\(s/n\)\s*:\s*$")

# Nome de cada ação como a operadora vê (botão e cabeçalho da prévia na saída).
ROTULOS_DAS_ACOES = {
    ACAO_CONFERIR: "Conferir antes de rodar",
    ACAO_SYNC: "Executar ciclo mensal",
    ACAO_CONCLUIR: "Registrar conclusão de curso",
    ACAO_RESETAR_TESTE: "Resetar ambiente de teste",
}

# Rótulos e marcas do painel do meio. Símbolos simples, que o Segoe UI desenha
# em qualquer Windows - nada de emoji, que vira quadradinho.
TITULO_PREVIA = " O que vai acontecer quando você clicar "
TITULO_ANDAMENTO = " O que está acontecendo agora "
TEXTO_VER_DETALHES = "Ver detalhes (caminhos completos)"
TEXTO_ESCONDER_DETALHES = "Esconder detalhes"
MARCA_PENDENTE = "○"       # circulo vazio
MARCA_ATUAL = "▶"          # triangulo
MARCA_FEITA = "✔"          # check
MARCA_ERRO = "✖"           # x
MARCA_NAO_EXECUTADA = "–"  # travessão: etapa que não chegou a rodar

FONTE_ETAPA = ("Segoe UI", 10)
FONTE_ETAPA_DESTAQUE = ("Segoe UI", 10, "bold")

# estado -> (marca, cor do texto, destaque?, complemento do texto)
VISUAL_DAS_ETAPAS = {
    ESTADO_PENDENTE: (MARCA_PENDENTE, "#666666", False, ""),
    ESTADO_ATUAL: (MARCA_ATUAL, "#000000", True, ""),
    ESTADO_FEITA: (MARCA_FEITA, "#1B7A3D", False, ""),
    ESTADO_NAO_EXECUTADA: (MARCA_NAO_EXECUTADA, "#8A8A8A", False, "   (não chegou a rodar)"),
    ESTADO_ERRO: (MARCA_ERRO, "#B00020", True, ""),
}

FONTE_STATUS = ("Segoe UI", 9)
FONTE_STATUS_DESTAQUE = ("Segoe UI", 9, "bold")

# nível do desfecho -> (cor do texto, cor de fundo ou None = a padrão, destaque?,
# tag da saída). O alerta é AMARELO de fundo: amarelo como cor de texto sobre o
# cinza do Windows não se lê.
VISUAL_DO_DESFECHO = {
    NIVEL_OK: ("#1B7A3D", None, True, "painel"),
    NIVEL_NEUTRO: ("#333333", None, True, "painel"),
    NIVEL_ALERTA: ("#000000", "#FFD54F", True, "aviso"),
    NIVEL_ERRO: ("#B00020", None, True, "erro"),
    NIVEL_ERRO_GRAVE: ("#FFFFFF", "#B00020", True, "erro"),
}

_CREATE_NO_WINDOW = 0x08000000


def caminho_do_logo() -> Path:
    """Onde o painel procura o logo: ``_Automação/logo.png``.

    É só largar o PNG nessa pasta (do lado dos ``.bat``) que ele aparece na
    faixa de marca. Nada a configurar.
    """
    return PASTA_AUTOMACAO / NOME_ARQUIVO_LOGO


def carregar_logo(janela: tk.Misc) -> "tk.PhotoImage | None":
    """O logo como ``PhotoImage``, ou ``None`` se não der.

    ``None`` é caso normal, não erro: hoje o arquivo nem está no repositório. O
    ``except`` largo é proposital - PNG corrompido, formato que este Tk não
    entende, arquivo sem permissão - nada disso pode impedir a operadora de
    fechar o mês. O único registro é uma linha de log.
    """
    arquivo = caminho_do_logo()
    if not arquivo.is_file():
        log.info(
            "Sem logo institucional em '%s' - o painel abre sem ele (é só largar o PNG "
            "nessa pasta).",
            arquivo,
        )
        return None
    try:
        imagem = tk.PhotoImage(master=janela, file=str(arquivo))
        # `subsample` só aceita inteiro: reduz pela metade, um terço, etc. Dá pra
        # encaixar o logo na faixa sem Pillow e sem distorcer a proporção.
        fator = max(1, -(-imagem.height() // ALTURA_MAX_LOGO))
        if fator > 1:
            imagem = imagem.subsample(fator, fator)
        return imagem
    except Exception as erro:  # noqa: BLE001
        log.warning(
            "Não consegui carregar o logo '%s' (%s). O painel segue normalmente sem ele.",
            arquivo, erro,
        )
        return None


def interpretador() -> Path:
    """O ``python.exe`` do venv.

    O ``.bat`` abre a janela com ``pythonw.exe`` (sem console). Os subprocessos
    precisam do ``python.exe`` normal — o ``pythonw`` não tem stdout pra ler.
    """
    executavel = Path(sys.executable)
    if executavel.name.lower() == "pythonw.exe":
        candidato = executavel.with_name("python.exe")
        if candidato.exists():
            return candidato
    return executavel


ARGUMENTO_SOMENTE_PRODUCAO = "--somente-producao"


def abre_somente_em_producao(argumentos: list[str]) -> bool:
    """O ``.bat`` da operadora abre o painel só com PRODUÇÃO (sem TESTE e reset)."""
    return ARGUMENTO_SOMENTE_PRODUCAO in argumentos


class Painel(tk.Tk):
    def __init__(self, somente_producao: bool = False) -> None:
        super().__init__()
        # Painel da operadora: só PRODUÇÃO, sem a escolha de modo e sem
        # "Resetar ambiente de teste" (pedido do responsável pelo processo, 05/10/2026). O painel
        # completo, com TESTE, fica pro time (tools/).
        self.somente_producao = somente_producao
        self.title("Automação de Treinamentos - Time de TI")
        self.geometry("1040x820")
        self.minsize(940, 680)

        self.modo = tk.StringVar(value=MODO_PRODUCAO if somente_producao else MODO_TESTE)
        self.arquivo_vigentes = tk.StringVar(value="")
        self.arquivo_obsoletos = tk.StringVar(value="")
        self.usar_mes_manual = tk.BooleanVar(value=False)
        self.mes_manual = tk.StringVar(value=str(datetime.now().month))
        self.ano_manual = tk.StringVar(value=str(datetime.now().year))

        self.fila: queue.Queue = queue.Queue()
        self.processo: subprocess.Popen | None = None
        self.cauda = ""  # fim da saída, pra reconhecer a pergunta s/n
        self.linha_parcial = ""  # sobra de linha entre dois pedaços lidos do processo
        # Qual ação está rodando agora. Decide se Cancelar/fechar são permitidos:
        # ação que escreve em planilha não pode ser interrompida pela janela.
        self.acao_em_curso: str | None = None
        self.com_vigentes = False
        # O Planos e Macs que o ciclo em curso cria (pra oferecer abrir no fim).
        self.arquivo_do_ciclo: Path | None = None
        self.rotulos_etapas: list[tk.Label] = []
        self.textos_etapas: list[str] = []
        self.etapa_atual = -1
        # Índices das etapas cujo marcador apareceu no log. É isto, e não "o
        # processo saiu com 0", que decide quais etapas ganham ✔ no fim.
        self.etapas_acesas: set[int] = set()
        # Todas as linhas que o processo escreveu nesta execução - é delas que
        # sai o desfecho ("nada foi alterado", restauração incompleta...).
        self.linhas_da_execucao: list[str] = []
        # Tudo o que muda O QUE vai rodar (modo, mês, arquivos). Fica travado
        # enquanto o subprocesso roda: trocar pra TESTE no meio de um ciclo de
        # PRODUÇÃO deixaria a tela verde com o processo escrevendo nas reais.
        self.controles_de_configuracao: list[tk.Widget] = []
        # Depois do destroy(), um `after` pendente estoura com "application has
        # been destroyed". A flag corta o laço da fila antes disso.
        self.fechando = False

        self._montar()
        self._atualizar_modo()
        self.after(100, self._drenar_fila)
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    # ------------------------------------------------------------------ layout
    def _estilo_da_marca(self) -> None:
        """Acabamentos na cor da marca: títulos das seções e barra de progresso.

        Detalhe fino de propósito. O que precisa saltar aos olhos nesta janela é
        se a operadora está em teste ou em produção, não a identidade visual.
        Tudo em ``try`` porque tema de ttk varia com a versão do Windows e não
        vale derrubar o painel por causa de cor de título.
        """
        try:
            estilo = ttk.Style(self)
            estilo.configure(
                "TLabelframe.Label", foreground=COR_MARCA_ESCURA, font=("Segoe UI", 9, "bold")
            )
            estilo.configure("Marca.Horizontal.TProgressbar", background=COR_MARCA)
        except tk.TclError as erro:
            log.warning("Não consegui aplicar o estilo visual (%s); seguindo com o padrão.", erro)

    def _faixa_de_marca(self) -> None:
        """Faixa institucional do topo: fundo branco, logo (se houver) e título.

        O turquesa da marca entra no texto e na régua de baixo, não no fundo —
        ver o comentário das cores: turquesa cheio encostado no letreiro verde
        do modo teste vira "dois verdes empilhados".
        """
        faixa = tk.Frame(self, bg=COR_FAIXA_MARCA)
        faixa.pack(fill="x")
        # A referência precisa viver enquanto a janela viver: PhotoImage some se
        # o Python coletar o objeto, e aí o logo vira um retângulo vazio.
        self.logo = carregar_logo(self)
        if self.logo is not None:
            tk.Label(faixa, image=self.logo, bg=COR_FAIXA_MARCA).pack(
                side="left", padx=(14, 10), pady=6
            )
        tk.Label(
            faixa,
            text="Automação de Treinamentos  •  Time de TI",
            font=("Segoe UI", 13, "bold"),
            bg=COR_FAIXA_MARCA,
            fg=COR_MARCA_ESCURA,
        ).pack(side="left", padx=(14 if self.logo is None else 0), pady=10)
        # Régua turquesa fechando a faixa. Fina de propósito: 3 pixels de marca
        # não disputam com o bloco enorme de cor do letreiro logo abaixo.
        tk.Frame(self, bg=COR_MARCA, height=3).pack(fill="x")

    def _montar(self) -> None:
        self._estilo_da_marca()
        self._faixa_de_marca()

        self.banner = tk.Label(self, text="", font=("Segoe UI", 20, "bold"), pady=14)
        self.banner.pack(fill="x")

        self.rotulo_base = tk.Label(self, text="", font=("Segoe UI", 10), pady=2)
        self.rotulo_base.pack(fill="x")

        numero = iter(range(1, 10))
        self.radios_modo: list[ttk.Radiobutton] = []
        if not self.somente_producao:
            seletor = ttk.LabelFrame(self, text=f" {next(numero)}. Onde eu vou mexer ")
            seletor.pack(fill="x", padx=12, pady=(10, 6))
            radio_teste = ttk.Radiobutton(
                seletor,
                text="Ambiente de TESTE — cópias das planilhas; posso errar à vontade",
                value=MODO_TESTE,
                variable=self.modo,
                command=self._atualizar_modo,
            )
            radio_teste.pack(anchor="w", padx=10, pady=(6, 0))
            radio_producao = ttk.Radiobutton(
                seletor,
                text="PRODUÇÃO — as planilhas REAIS do time, as que valem",
                value=MODO_PRODUCAO,
                variable=self.modo,
                command=self._atualizar_modo,
            )
            radio_producao.pack(anchor="w", padx=10, pady=(0, 8))
            self.radios_modo = [radio_teste, radio_producao]
            self.controles_de_configuracao += self.radios_modo

        arquivos = ttk.LabelFrame(
            self, text=f" {next(numero)}. Os dois arquivos que chegaram por e-mail "
        )
        arquivos.pack(fill="x", padx=12, pady=6)
        self.rotulo_vigentes = self._linha_de_arquivo(
            arquivos, "POPs vigentes (aprovados e efetivos):", self._escolher_vigentes,
            self._limpar_vigentes,
        )
        self.rotulo_obsoletos = self._linha_de_arquivo(
            arquivos, "POPs obsoletos:", self._escolher_obsoletos, self._limpar_obsoletos,
        )

        mes = ttk.LabelFrame(self, text=f" {next(numero)}. Mês dos POPs (o Planos e Macs que vai ser criado) ")
        mes.pack(fill="x", padx=12, pady=6)
        linha = ttk.Frame(mes)
        linha.pack(fill="x", padx=10, pady=6)
        self.check_mes = ttk.Checkbutton(
            linha,
            text="Escolher o mês (senão, usa o mês passado - os POPs chegam no começo do mês seguinte)",
            variable=self.usar_mes_manual,
            command=self._atualizar_plano,
        )
        self.check_mes.pack(side="left")
        ttk.Label(linha, text="   mês:").pack(side="left")
        self.spin_mes = ttk.Spinbox(
            linha, from_=1, to=12, width=4, textvariable=self.mes_manual,
            command=self._atualizar_plano,
        )
        self.spin_mes.pack(side="left")
        ttk.Label(linha, text="  ano:").pack(side="left")
        self.spin_ano = ttk.Spinbox(
            linha, from_=2020, to=2100, width=6, textvariable=self.ano_manual,
            command=self._atualizar_plano,
        )
        self.spin_ano.pack(side="left")
        self.controles_de_configuracao += [self.check_mes, self.spin_mes, self.spin_ano]

        acoes = ttk.LabelFrame(self, text=f" {next(numero)}. O que fazer ")
        acoes.pack(fill="x", padx=12, pady=6)
        grade = ttk.Frame(acoes)
        grade.pack(fill="x", padx=10, pady=8)
        self.botoes: dict[str, ttk.Button] = {}
        botoes = (
            (ACAO_CONFERIR, f"{ROTULOS_DAS_ACOES[ACAO_CONFERIR]}\n(não altera nada)"),
            (ACAO_SYNC, ROTULOS_DAS_ACOES[ACAO_SYNC]),
            (ACAO_CONCLUIR, ROTULOS_DAS_ACOES[ACAO_CONCLUIR]),
            (ACAO_RESETAR_TESTE, ROTULOS_DAS_ACOES[ACAO_RESETAR_TESTE]),
        )
        if self.somente_producao:
            botoes = tuple(b for b in botoes if b[0] != ACAO_RESETAR_TESTE)
        for coluna, (acao, rotulo) in enumerate(botoes):
            botao = ttk.Button(grade, text=rotulo, command=lambda a=acao: self._executar(a))
            botao.grid(row=0, column=coluna, sticky="ew", padx=4, ipady=8)
            grade.columnconfigure(coluna, weight=1)
            self.botoes[acao] = botao

        # Um painel só, com dois estados: ANTES de rodar mostra a prévia
        # ("o que vai acontecer"); DURANTE e DEPOIS mostra o andamento ("o que
        # está acontecendo agora"). São as duas perguntas que a operadora faz, e
        # nunca ao mesmo tempo - dividir o espaço entre elas deixaria as duas
        # apertadas.
        self.painel_meio = ttk.LabelFrame(self, text=TITULO_PREVIA)
        self.painel_meio.pack(fill="x", padx=12, pady=6)

        self.quadro_previa = ttk.Frame(self.painel_meio)
        self.rotulo_previa = tk.Label(
            self.quadro_previa, text="", justify="left", anchor="w",
            font=("Segoe UI", 10), wraplength=960,
        )
        self.rotulo_previa.pack(fill="x", padx=10, pady=(6, 2))
        self.botao_detalhes = ttk.Button(
            self.quadro_previa, text=TEXTO_VER_DETALHES, command=self._alternar_detalhes
        )
        self.botao_detalhes.pack(anchor="w", padx=10, pady=(0, 6))
        # Recolhido por padrão: caminho absoluto é material de diagnóstico, não
        # informação de operação.
        self.texto_detalhes = tk.Text(
            self.quadro_previa, height=7, wrap="none", font=("Consolas", 8),
            bg="#F6F6F6", relief="flat",
        )

        self.quadro_andamento = ttk.Frame(self.painel_meio)
        self.quadro_etapas = ttk.Frame(self.quadro_andamento)
        self.quadro_etapas.pack(fill="x", padx=10, pady=(6, 2))
        self.rotulo_mexendo = tk.Label(
            self.quadro_andamento, text="", justify="left", anchor="w",
            font=("Segoe UI", 9, "italic"), fg=COR_MARCA_ESCURA,
        )
        self.rotulo_mexendo.pack(fill="x", padx=10, pady=(0, 6))

        saida = ttk.LabelFrame(self, text=" Detalhes técnicos (selecione e copie se der erro) ")
        saida.pack(fill="both", expand=True, padx=12, pady=(6, 4))
        quadro = ttk.Frame(saida)
        quadro.pack(fill="both", expand=True, padx=8, pady=6)
        self.saida = tk.Text(
            quadro, wrap="word", height=7, bg=COR_FUNDO_SAIDA, fg=COR_TEXTO_SAIDA,
            insertbackground=COR_TEXTO_SAIDA, font=("Consolas", 8),
        )
        barra = ttk.Scrollbar(quadro, orient="vertical", command=self.saida.yview)
        self.saida.configure(yscrollcommand=barra.set)
        self.saida.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.saida.tag_configure("erro", foreground="#FF6B6B")
        self.saida.tag_configure("aviso", foreground="#FFC857")
        self.saida.tag_configure("painel", foreground="#7FD1FF")

        # Régua fina na cor da marca fechando o corpo da janela - acabamento,
        # sem disputar atenção com o letreiro de modo.
        tk.Frame(self, bg=COR_MARCA, height=3).pack(fill="x", padx=12, pady=(4, 6))

        rodape = ttk.Frame(self)
        rodape.pack(fill="x", padx=12, pady=(0, 10))
        self.progresso = ttk.Progressbar(
            rodape, mode="indeterminate", length=180,
            style="Marca.Horizontal.TProgressbar",
        )
        self.progresso.pack(side="left")
        # tk.Label (e não ttk) porque o desfecho pinta o FUNDO: amarelo no
        # alerta, vermelho no erro grave. O ttk ignora `bg` em vários temas.
        self.rotulo_status = tk.Label(rodape, text="Parado.", font=FONTE_STATUS, padx=6)
        self.rotulo_status.pack(side="left", padx=10)
        self._fundo_status = self.rotulo_status.cget("bg")
        self._texto_status = self.rotulo_status.cget("fg")
        ttk.Button(rodape, text="Copiar saída", command=self._copiar_saida).pack(side="right")
        ttk.Button(rodape, text="Limpar saída", command=self._limpar_saida).pack(
            side="right", padx=6
        )
        self.botao_cancelar = ttk.Button(
            rodape, text="Cancelar execução", command=self._cancelar, state="disabled"
        )
        self.botao_cancelar.pack(side="right", padx=6)

        resposta = ttk.Frame(self)
        resposta.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Label(resposta, text="Responder ao programa:").pack(side="left")
        self.entrada_resposta = ttk.Entry(resposta, width=20, state="disabled")
        self.entrada_resposta.pack(side="left", padx=6)
        self.entrada_resposta.bind("<Return>", lambda _e: self._enviar_resposta())
        self.botao_responder = ttk.Button(
            resposta, text="Enviar", command=self._enviar_resposta, state="disabled"
        )
        self.botao_responder.pack(side="left")

    def _linha_de_arquivo(self, pai, titulo, ao_escolher, ao_limpar) -> tk.Label:
        linha = ttk.Frame(pai)
        linha.pack(fill="x", padx=10, pady=5)
        ttk.Label(linha, text=titulo, width=36).pack(side="left")
        escolher = ttk.Button(linha, text="Escolher...", command=ao_escolher)
        escolher.pack(side="left")
        tirar = ttk.Button(linha, text="Tirar", command=ao_limpar)
        tirar.pack(side="left", padx=4)
        self.controles_de_configuracao += [escolher, tirar]
        rotulo = tk.Label(linha, text="(nenhum arquivo escolhido)", anchor="w", justify="left")
        rotulo.pack(side="left", fill="x", expand=True, padx=8)
        return rotulo

    # ------------------------------------------------------------------ estado
    def _atualizar_modo(self) -> None:
        teste = self.modo.get() == MODO_TESTE
        cor = COR_TESTE_FUNDO if teste else COR_PRODUCAO_FUNDO
        texto = COR_TEXTO_TESTE if teste else COR_TEXTO_PRODUCAO
        self.banner.configure(text=texto_do_banner(self.modo.get()), bg=cor, fg=texto)
        self.configure(bg=cor)
        base = base_do_modo(self.modo.get())
        self.rotulo_base.configure(
            text=f"Pasta em uso:  {base}",
            bg=cor,
            fg=texto,
        )
        self._atualizar_plano()

    def _meses_escolhidos(self) -> tuple[int | None, int | None]:
        if not self.usar_mes_manual.get():
            return None, None
        try:
            return int(self.mes_manual.get()), int(self.ano_manual.get())
        except ValueError:
            return None, None

    def _meses(self) -> dict | None:
        from treinamentos_its.sync import obter_meses

        mes, ano = self._meses_escolhidos()
        try:
            return obter_meses(mes, ano)
        except ValueError:
            return None

    def _atualizar_plano(self) -> None:
        """Atualiza a prévia do ciclo mensal e o texto escondido de detalhes.

        Em linguagem de gente: mês e nome de arquivo, nada de caminho. Os
        caminhos completos ficam no "ver detalhes", recolhido. A prévia das
        OUTRAS ações sai na saída, no começo de cada execução.
        """
        meses = self._meses()
        vigentes = self.arquivo_vigentes.get() or None
        obsoletos = self.arquivo_obsoletos.get() or None
        frases = resumo_humano(
            ACAO_SYNC, meses, modo=self.modo.get(), vigentes=vigentes, obsoletos=obsoletos
        )
        self.rotulo_previa.configure(
            text=f"Se você clicar em '{ROTULOS_DAS_ACOES[ACAO_SYNC]}' agora:\n"
            + "\n".join(f"   • {frase}" for frase in frases)
        )

        le, escreve = plano_de_arquivos(
            ACAO_SYNC, self.modo.get(), meses, vigentes=vigentes, obsoletos=obsoletos
        )
        linhas = ["LÊ:"] + [f"   {item}" for item in le or ["(nada)"]]
        linhas += ["", "CRIA / ALTERA:"] + [f"   {item}" for item in escreve or ["(nada)"]]
        self.texto_detalhes.configure(state="normal")
        self.texto_detalhes.delete("1.0", "end")
        self.texto_detalhes.insert("1.0", "\n".join(linhas))
        self.texto_detalhes.configure(state="disabled")
        self._mostrar_previa()

    def _alternar_detalhes(self) -> None:
        if self.texto_detalhes.winfo_ismapped():
            self.texto_detalhes.pack_forget()
            self.botao_detalhes.configure(text=TEXTO_VER_DETALHES)
        else:
            self.texto_detalhes.pack(fill="x", padx=10, pady=(0, 8))
            self.botao_detalhes.configure(text=TEXTO_ESCONDER_DETALHES)

    # ------------------------------------------------------------- andamento
    def _mostrar_previa(self) -> None:
        """Volta o painel do meio pra prévia (nada rodando)."""
        if self.processo is not None:
            return  # no meio de uma execução, o andamento manda
        self.quadro_andamento.pack_forget()
        self.quadro_previa.pack(fill="x")
        self.painel_meio.configure(text=TITULO_PREVIA)

    def _iniciar_andamento(self, acao: str, com_vigentes: bool) -> None:
        """Troca a prévia pela lista de etapas daquela ação."""
        self.acao_em_curso = acao
        self.com_vigentes = com_vigentes
        self.etapa_atual = -1
        self.etapas_acesas = set()
        self.linhas_da_execucao = []
        self.linha_parcial = ""
        for rotulo in self.rotulos_etapas:
            rotulo.destroy()
        self.rotulos_etapas = []
        self.textos_etapas = list(etapas_da_acao(acao, com_vigentes))
        for texto in self.textos_etapas:
            rotulo = tk.Label(
                self.quadro_etapas, text=f"{MARCA_PENDENTE}  {texto}",
                anchor="w", justify="left", font=FONTE_ETAPA, fg="#666666",
            )
            rotulo.pack(fill="x")
            self.rotulos_etapas.append(rotulo)
        if not self.rotulos_etapas:
            rotulo = tk.Label(
                self.quadro_etapas, text=f"{MARCA_ATUAL}  Em andamento...",
                anchor="w", font=FONTE_ETAPA,
            )
            rotulo.pack(fill="x")
            self.rotulos_etapas.append(rotulo)
        self.rotulo_mexendo.configure(text="")
        self.quadro_previa.pack_forget()
        self.quadro_andamento.pack(fill="x")
        self.painel_meio.configure(text=TITULO_ANDAMENTO)

    def _pintar_etapas(self, estados: list[str]) -> None:
        """Pinta cada etapa conforme ``plano_execucao.estados_das_etapas``."""
        for rotulo, texto, estado in zip(self.rotulos_etapas, self.textos_etapas, estados):
            marca, cor, destaque, complemento = VISUAL_DAS_ETAPAS[estado]
            rotulo.configure(
                text=f"{marca}  {texto}{complemento}",
                fg=cor,
                font=FONTE_ETAPA_DESTAQUE if destaque else FONTE_ETAPA,
            )

    def _processar_linha(self, linha: str) -> None:
        """Guarda UMA linha completa do programa e atualiza o andamento."""
        self.linhas_da_execucao.append(linha)
        self._andamento_da_linha(linha)

    def _andamento_da_linha(self, linha: str) -> None:
        """Lê UMA linha do programa e atualiza a etapa/arquivo em curso."""
        if self.acao_em_curso is None:
            return
        indice = indice_da_etapa(self.acao_em_curso, linha, self.com_vigentes)
        if indice is not None:
            self.etapas_acesas.add(indice)
            self.etapa_atual = max(self.etapa_atual, indice)
            self._pintar_etapas(
                estados_das_etapas(len(self.textos_etapas), self.etapas_acesas)
            )
        nome = arquivo_da_linha(linha)
        if nome:
            self.rotulo_mexendo.configure(text=f"mexendo agora em:  {nome}")

    # ---------------------------------------------------------------- arquivos
    def _pasta_inicial(self) -> str:
        base = base_do_modo(self.modo.get())
        for candidata in (base / "Planos e Macs" / "POPs do Sistema", base / "Planos e Macs", base):
            if candidata.is_dir():
                return str(candidata)
        return str(PASTA_AUTOMACAO)

    def _escolher(self, titulo: str) -> str:
        return filedialog.askopenfilename(
            title=titulo,
            initialdir=self._pasta_inicial(),
            filetypes=[("Planilhas do Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
        )

    def _descrever_arquivo(self, caminho: str) -> str:
        arquivo = Path(caminho)
        mes, ano = inferir_mes_ano(arquivo.name)
        try:
            momento = datetime.fromtimestamp(arquivo.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
        except OSError:
            momento = "?"
        return f"{arquivo.name}\n{descrever_mes_inferido(mes, ano)}  •  modificado em {momento}"

    def _pintar_rotulo(self, rotulo: tk.Label, caminho: str) -> None:
        if not caminho:
            rotulo.configure(text="(nenhum arquivo escolhido)", fg="black")
            return
        mes, _ano = inferir_mes_ano(Path(caminho).name)
        rotulo.configure(
            text=self._descrever_arquivo(caminho),
            # Mês não identificado não é erro, mas merece destaque: é justamente o
            # caso em que a operadora não tem como saber se pegou o anexo certo.
            fg="black" if mes else "#A05A00",
        )

    def _escolher_vigentes(self) -> None:
        caminho = self._escolher("Escolha o export de POPs VIGENTES (aprovados e efetivos)")
        if caminho:
            self.arquivo_vigentes.set(caminho)
            self._pintar_rotulo(self.rotulo_vigentes, caminho)
            self._atualizar_plano()

    def _escolher_obsoletos(self) -> None:
        caminho = self._escolher("Escolha o export de POPs OBSOLETOS")
        if caminho:
            self.arquivo_obsoletos.set(caminho)
            self._pintar_rotulo(self.rotulo_obsoletos, caminho)
            self._atualizar_plano()

    def _limpar_vigentes(self) -> None:
        self.arquivo_vigentes.set("")
        self._pintar_rotulo(self.rotulo_vigentes, "")
        self._atualizar_plano()

    def _limpar_obsoletos(self) -> None:
        self.arquivo_obsoletos.set("")
        self._pintar_rotulo(self.rotulo_obsoletos, "")
        self._atualizar_plano()

    # ---------------------------------------------------------------- execução
    def _executar(self, acao: str) -> None:
        if self.processo is not None:
            messagebox.showinfo(
                "Já tem coisa rodando",
                "Espere a execução atual terminar (ou clique em 'Cancelar execução').",
            )
            return

        # O modo é lido UMA vez: o que foi confirmado é o que roda.
        modo = self.modo.get()

        # O "mês do ciclo" só vale pro ciclo e pro conferir. No concluir o mês é o
        # da data da conclusão (ver `plano_execucao.ACOES_COM_MES`) - então um mês
        # inválido no seletor não pode travar o concluir, nem ir parar nele.
        mes, ano = None, None
        meses = None
        if acao in ACOES_COM_MES:
            mes, ano = self._meses_escolhidos()
            if self.usar_mes_manual.get():
                problema = problema_no_mes_escolhido(mes, ano)
                if problema:
                    messagebox.showerror("Mês do ciclo inválido", problema)
                    return
            meses = self._meses()
            if meses is None:
                messagebox.showerror(
                    "Mês do ciclo inválido",
                    "Não consegui montar o mês do ciclo com esse mês e ano.",
                )
                return

        vigentes = self.arquivo_vigentes.get() or None
        obsoletos = self.arquivo_obsoletos.get() or None

        motivo = motivo_para_nao_rodar(acao, modo, vigentes)
        if motivo:
            messagebox.showwarning("Falta o arquivo de POPs vigentes", motivo)
            return

        if acao in (ACAO_SYNC, ACAO_CONFERIR):
            usados = [vigentes] + ([obsoletos] if acao == ACAO_CONFERIR else [])
            for caminho in usados:
                if caminho and not Path(caminho).exists():
                    messagebox.showerror("Arquivo sumiu", f"Não encontrei mais:\n{caminho}")
                    return
            # Export de outro mês: o painel só mostrava o mês ao lado do arquivo
            # e rodava sem perguntar (QA de 05/10/2026).
            exports = {}
            if vigentes:
                exports["vigentes"] = inferir_mes_ano(Path(vigentes).name)
            if obsoletos and acao == ACAO_CONFERIR:
                exports["obsoletos"] = inferir_mes_ano(Path(obsoletos).name)
            avisos = avisos_de_mes_dos_exports(meses, exports)
            if avisos and not messagebox.askyesno(
                "Export de outro mês?",
                "\n\n".join(avisos) + "\n\nSeguir mesmo assim com esse(s) arquivo(s)?",
                default="no",
            ):
                return

        extras: dict = {}
        if acao == ACAO_CONCLUIR:
            dados = self._perguntar_conclusao()
            if dados is None:
                return
            extras = dados

        frases = resumo_humano(
            acao, meses, modo=modo, vigentes=vigentes, obsoletos=obsoletos, **extras
        )
        if not self._confirmar(acao, modo, frases):
            return

        argumentos = argumentos_do_comando(
            acao,
            modo,
            mes=mes,
            ano=ano,
            vigentes=vigentes,
            obsoletos=obsoletos,
            **extras,
        )
        # A prévia da AÇÃO clicada vai pra saída antes de rodar. O quadro do meio
        # só mostra a do ciclo; sem isto, a do conferir/concluir/resetar nunca
        # chegava na operadora (e as que não escrevem nem passam por confirmação).
        self._escrever_previa(acao, modo, frases)
        # A etapa de colagem só existe quando há export de vigentes - no conferir
        # também: sem o arquivo ele confere a Matriz do mês anterior como está.
        self._iniciar_andamento(acao, bool(vigentes) and acao in (ACAO_SYNC, ACAO_CONFERIR))
        self.arquivo_do_ciclo = (
            arquivo_criado_pelo_ciclo(modo, meses) if acao == ACAO_SYNC else None
        )
        self._rodar(argumentos, acao)

    def _escrever_previa(self, acao: str, modo: str, frases: list[str]) -> None:
        nome = ROTULOS_DAS_ACOES.get(acao, acao)
        if acao in ACOES_COM_TESTE:
            onde = "TESTE (cópias do sandbox)" if modo == MODO_TESTE else "PRODUÇÃO (planilhas reais)"
        else:
            onde = "só o ambiente de teste"
        texto = f"\n>>> {nome}  —  {onde}\nO que vai acontecer:\n"
        texto += "".join(f"   • {frase}\n" for frase in frases)
        self._escrever(texto, "painel")

    def _confirmar(self, acao: str, modo: str, frases: list[str]) -> bool:
        # A caixa de confirmação fala a MESMA língua da prévia. Era aqui que o
        # muro de caminhos aparecia; quem quiser caminho abre o "ver detalhes".
        plano = "\n".join(f"• {frase}" for frase in frases)
        if not acao_escreve_em_arquivo(acao):
            return True

        if acao == ACAO_RESETAR_TESTE:
            # Mexe SÓ no sandbox, mesmo com a tela em produção. Mostrar aqui o
            # aviso vermelho de "planilhas reais" seria mentira - e aviso que
            # mente é aviso que a pessoa aprende a ignorar.
            return messagebox.askokcancel(
                "Resetar o ambiente de teste",
                f"{plano}\n\nIsto NÃO toca nas planilhas reais: elas só são lidas para "
                "recriar as cópias. Tudo o que estiver no sandbox se perde.",
            )

        if modo != MODO_TESTE:
            aviso = (
                "ATENÇÃO: isto vai alterar as PLANILHAS REAIS do time.\n\n"
                f"{plano}\n\n"
                "Feche todas as janelas do Excel antes de continuar.\n\n"
                "Confirma?"
            )
            return messagebox.askyesno("PRODUÇÃO — confirmar alteração real", aviso, default="no")

        return messagebox.askokcancel(
            "Ambiente de teste — confirmar",
            f"Isto mexe só nas cópias do sandbox.\n\n{plano}\n\n"
            "Feche todas as janelas do Excel antes de continuar.",
        )

    def _perguntar_conclusao(self) -> dict | None:
        janela = tk.Toplevel(self)
        janela.title("Registrar conclusão de curso")
        janela.transient(self)
        janela.grab_set()
        campos: dict[str, tk.StringVar] = {
            "nome": tk.StringVar(),
            "codigo": tk.StringVar(),
            "data": tk.StringVar(value=datetime.now().strftime("%d/%m/%Y")),
        }
        rotulos = (
            ("nome", "Nome (igual ao da planilha):"),
            ("codigo", "Código do curso:"),
            ("data", "Data da conclusão (dd/mm/aaaa):"),
        )
        for numero, (chave, rotulo) in enumerate(rotulos):
            ttk.Label(janela, text=rotulo).grid(row=numero, column=0, sticky="w", padx=10, pady=6)
            ttk.Entry(janela, textvariable=campos[chave], width=42).grid(
                row=numero, column=1, padx=10, pady=6
            )
        resultado: dict = {}

        def confirmar() -> None:
            if not campos["nome"].get().strip() or not campos["codigo"].get().strip():
                messagebox.showerror("Faltou preencher", "Nome e código são obrigatórios.", parent=janela)
                return
            try:
                datetime.strptime(campos["data"].get().strip(), "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Data inválida", "Use dd/mm/aaaa.", parent=janela)
                return
            resultado.update(
                nome=campos["nome"].get().strip(),
                codigo=campos["codigo"].get().strip(),
                data=campos["data"].get().strip(),
            )
            janela.destroy()

        ttk.Button(janela, text="Registrar", command=confirmar).grid(row=3, column=1, sticky="e", padx=10, pady=10)
        ttk.Button(janela, text="Cancelar", command=janela.destroy).grid(row=3, column=0, sticky="w", padx=10, pady=10)
        self.wait_window(janela)
        return resultado or None

    def _rodar(self, argumentos: list[str], acao: str) -> None:
        comando = [str(interpretador()), "-u", "main.py", *argumentos]
        self._escrever(f"\n$ python main.py {' '.join(argumentos)}\n", "painel")
        ambiente = {
            **os.environ,
            "PYTHONUTF8": "1",
            "PYTHONIOENCODING": "utf-8",
        }
        self.cauda = ""
        # Trava ANTES de abrir o processo: não sobra instante em que o modo possa
        # ser trocado com o comando já montado.
        self._travar_botoes(True)
        self._pintar_status(f"Rodando '{acao}'... (o Excel abre por baixo)")
        try:
            self.processo = subprocess.Popen(
                comando,
                cwd=str(PASTA_AUTOMACAO),
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=ambiente,
                creationflags=_CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
        except OSError as erro:
            self._escrever(f"Não consegui iniciar o programa: {erro}\n", "erro")
            # Fecha a execução como erro: destrava a janela e não deixa o
            # andamento parado numa ação que nem começou.
            self._terminou(-1)
            return

        self.progresso.start(12)
        threading.Thread(target=self._ler_saida, args=(self.processo,), daemon=True).start()

    def _ler_saida(self, processo: subprocess.Popen) -> None:
        """Roda numa thread: lê a saída em pedaços e joga na fila da interface."""
        decodificador = codecs.getincrementaldecoder("utf-8")(errors="replace")
        try:
            while True:
                pedaco = processo.stdout.read1(4096) if processo.stdout else b""
                if not pedaco:
                    break
                self.fila.put(("texto", decodificador.decode(pedaco)))
        except Exception as erro:  # noqa: BLE001
            self.fila.put(("texto", f"\n[erro lendo a saída: {erro}]\n"))
        finally:
            resto = decodificador.decode(b"", final=True)
            if resto:
                self.fila.put(("texto", resto))
            self.fila.put(("fim", processo.wait()))

    def _drenar_fila(self) -> None:
        if self.fechando:
            return
        try:
            while True:
                tipo, carga = self.fila.get_nowait()
                if tipo == "texto":
                    self._escrever(carga)
                    self.cauda = (self.cauda + carga)[-200:]
                    # O andamento precisa de LINHAS inteiras; a leitura vem em
                    # pedaços, então a sobra fica guardada até fechar a linha.
                    self.linha_parcial += carga
                    *completas, self.linha_parcial = self.linha_parcial.split("\n")
                    for linha in completas:
                        self._processar_linha(linha)
                    if PADRAO_PERGUNTA.search(self.cauda):
                        self.cauda = ""
                        self.after(50, self._responder_pergunta)
                elif tipo == "fim":
                    self._terminou(carga)
        except queue.Empty:
            pass
        self.after(100, self._drenar_fila)

    def _responder_pergunta(self) -> None:
        """O programa fez uma pergunta s/n no terminal; leva pra uma caixa."""
        linhas = self.saida.get("1.0", "end-1c").rstrip().splitlines()
        quer = messagebox.askyesno(
            "O programa está perguntando",
            texto_da_pergunta(linhas),
            default="no",
        )
        self._mandar_para_o_processo("s" if quer else "n")

    def _enviar_resposta(self) -> None:
        texto = self.entrada_resposta.get().strip()
        self.entrada_resposta.delete(0, "end")
        self._mandar_para_o_processo(texto)

    def _mandar_para_o_processo(self, texto: str) -> None:
        if self.processo is None or self.processo.stdin is None:
            return
        try:
            self.processo.stdin.write(f"{texto}\n".encode("utf-8"))
            self.processo.stdin.flush()
            self._escrever(f"{texto}\n", "painel")
        except OSError as erro:
            self._escrever(f"[não consegui responder ao programa: {erro}]\n", "erro")

    def _cancelar(self) -> None:
        if self.processo is None:
            return
        if acao_escreve_em_arquivo(self.acao_em_curso or ""):
            # Defesa: o botão já fica desabilitado nesse caso.
            messagebox.showwarning(
                "Não dá para cancelar agora",
                "Esta execução está alterando planilha. Interromper pela metade "
                "deixaria arquivo inconsistente, e as cópias de segurança não "
                "ajudariam: quem restaura é o próprio programa, e ele morreria "
                "junto.\n\nEspere terminar - se der erro, ele desfaz sozinho.",
            )
            return
        if not messagebox.askyesno(
            "Cancelar a conferência?",
            "Esta ação só lê as planilhas, não altera nada - dá para "
            "interromper com segurança.\n\nCancelar?",
            default="no",
        ):
            return
        try:
            self.processo.terminate()
            self._escrever("\n[execução cancelada pela operadora]\n", "erro")
        except OSError as erro:
            self._escrever(f"[não consegui cancelar: {erro}]\n", "erro")

    def _terminou(self, codigo: int) -> None:
        acao = self.acao_em_curso or ""
        # A sobra de linha é processada ANTES de soltar a ação: com a ação já
        # zerada, `_andamento_da_linha` ignorava a última linha (a que não
        # termina em quebra de linha) e a etapa dela nunca acendia.
        if self.linha_parcial:
            self._processar_linha(self.linha_parcial)
            self.linha_parcial = ""
        self.processo = None
        self.acao_em_curso = None
        self.progresso.stop()
        self._travar_botoes(False)

        # Exit 0 NÃO é "fiz o trabalho": guarda respondida "n" e conclusão
        # parcial também saem com 0. Quem traduz é o `desfecho_da_execucao`.
        texto, nivel = desfecho_da_execucao(acao, codigo, self.linhas_da_execucao)
        sucesso = nivel in NIVEIS_SEM_ERRO
        if self.textos_etapas:
            self._pintar_etapas(
                estados_das_etapas(
                    len(self.textos_etapas), self.etapas_acesas, terminou=True, sucesso=sucesso
                )
            )
        elif self.rotulos_etapas:
            marca = MARCA_FEITA if sucesso else MARCA_ERRO
            self.rotulos_etapas[0].configure(text=f"{marca}  {texto}")

        if sucesso:
            self.rotulo_mexendo.configure(text=texto)
        elif nivel == NIVEL_ERRO_GRAVE:
            self.rotulo_mexendo.configure(
                text=f"{texto}. NÃO rode de novo: copie os detalhes técnicos abaixo e "
                "mande pro time agora."
            )
        else:
            self.rotulo_mexendo.configure(
                text="Parou aqui. Copie os detalhes técnicos abaixo e mande pro time."
            )

        _cor, _fundo, _destaque, tag = VISUAL_DO_DESFECHO[nivel]
        final = texto.rstrip(".")
        if nivel == NIVEL_ERRO:
            final += " — leia as mensagens acima"
        self._escrever(f"\n=== {final} ===\n", tag)
        self._pintar_status(texto, nivel)
        if oferecer_abrir_o_arquivo(acao, nivel, self.linhas_da_execucao):
            self.after(300, self._oferecer_abrir, self.arquivo_do_ciclo)

    def _oferecer_abrir(self, arquivo: Path | None) -> None:
        """Pergunta se abre o Planos e Macs que o ciclo acabou de gravar."""
        if arquivo is None or not arquivo.exists() or self.fechando:
            return
        if messagebox.askyesno(
            "Ciclo concluído",
            f"Pronto! O {arquivo.stem} está salvo em:\n{arquivo.parent}\n\n"
            "Quer abrir agora para conferir?",
            default="yes",
        ):
            try:
                os.startfile(arquivo)  # noqa: S606 - abre no Excel da operadora
            except OSError as erro:
                messagebox.showerror("Não consegui abrir", f"{arquivo}\n\n{erro}")

    def _pintar_status(self, texto: str, nivel: str | None = None) -> None:
        """O texto do rodapé; com ``nivel``, na cor do desfecho."""
        if nivel is None:
            self.rotulo_status.configure(
                text=texto, fg=self._texto_status, bg=self._fundo_status, font=FONTE_STATUS
            )
            return
        cor, fundo, destaque, _tag = VISUAL_DO_DESFECHO[nivel]
        self.rotulo_status.configure(
            text=texto,
            fg=cor,
            bg=fundo or self._fundo_status,
            font=FONTE_STATUS_DESTAQUE if destaque else FONTE_STATUS,
        )

    def _travar_botoes(self, rodando: bool) -> None:
        estado = "disabled" if rodando else "normal"
        for botao in self.botoes.values():
            botao.configure(state=estado)
        # Modo, mês e arquivos também: são eles que decidem O QUE está rodando.
        # Destravados, a operadora trocava pra TESTE no meio de um ciclo de
        # PRODUÇÃO e a tela ficava verde enquanto o processo escrevia nas reais.
        for controle in self.controles_de_configuracao:
            controle.configure(state=estado)
        # Cancelar só existe para ação que NÃO escreve (o `conferir`). Matar o
        # processo no meio de uma escrita pula o `finally`, deixa o Excel aberto
        # e, com o AutoSave do OneDrive, persiste a alteração parcial - ou seja,
        # o cancelamento destruiria justamente o rollback que deveria salvar.
        pode_cancelar = rodando and not acao_escreve_em_arquivo(self.acao_em_curso or "")
        self.botao_cancelar.configure(state="normal" if pode_cancelar else "disabled")
        self.entrada_resposta.configure(state="normal" if rodando else "disabled")
        self.botao_responder.configure(state="normal" if rodando else "disabled")

    # -------------------------------------------------------------------- saída
    def _escrever(self, texto: str, tag: str | None = None) -> None:
        primeira = int(self.saida.index("end-1c").split(".")[0])
        self.saida.insert("end", texto)
        if tag:
            ultima = int(self.saida.index("end-1c").split(".")[0])
            self.saida.tag_add(tag, f"{primeira}.0", f"{ultima}.end")
        else:
            self._colorir(primeira)
        self.saida.see("end")

    def _colorir(self, primeira_linha: int) -> None:
        ultima = int(self.saida.index("end-1c").split(".")[0])
        for numero in range(primeira_linha, ultima + 1):
            conteudo = self.saida.get(f"{numero}.0", f"{numero}.end")
            if " ERROR " in conteudo or conteudo.startswith("Traceback"):
                self.saida.tag_add("erro", f"{numero}.0", f"{numero}.end")
            elif " WARNING " in conteudo:
                self.saida.tag_add("aviso", f"{numero}.0", f"{numero}.end")

    def _copiar_saida(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.saida.get("1.0", "end-1c"))
        self._pintar_status("Saída copiada para a área de transferência.")

    def _limpar_saida(self) -> None:
        self.saida.delete("1.0", "end")

    def _ao_fechar(self) -> None:
        if self.processo is not None:
            if acao_escreve_em_arquivo(self.acao_em_curso or ""):
                # Fechar a janela mataria o subprocesso no meio de uma escrita.
                # Não é escolha da operadora: ela não tem como saber em que ponto
                # do arquivo o programa está.
                messagebox.showwarning(
                    "Espere terminar",
                    "Tem uma execução alterando planilha agora. Fechar a janela "
                    "interromperia pela metade e impediria o programa de desfazer "
                    "o que já fez.\n\nDeixe terminar - falta pouco. Se der erro, "
                    "ele restaura sozinho.",
                )
                return
            if not messagebox.askyesno(
                "Ainda está rodando",
                "Tem uma conferência em andamento. Ela só lê as planilhas, então "
                "dá para fechar sem risco.\n\nFechar?",
                default="no",
            ):
                return
            try:
                self.processo.terminate()
            except OSError:
                pass
        self.fechando = True
        self.destroy()


def main(argumentos: list[str] | None = None) -> None:
    somente_producao = abre_somente_em_producao(
        sys.argv[1:] if argumentos is None else argumentos
    )
    painel = Painel(somente_producao=somente_producao)
    if somente_producao:
        painel._escrever(
            "Painel aberto em PRODUÇÃO: o que roda aqui mexe nas planilhas do time, "
            f"em {base_do_modo(MODO_PRODUCAO)}.\n"
            "Sempre: 'Conferir antes de rodar' primeiro, e leia os avisos.\n",
            "painel",
        )
    else:
        painel._escrever(
            "Painel aberto. O ambiente começa em TESTE de propósito: o que roda aqui "
            f"mexe só nas cópias em {PASTA_SANDBOX}.\n"
            "Para valer nas planilhas do time, troque para PRODUÇÃO lá em cima.\n",
            "painel",
        )
    painel.mainloop()


if __name__ == "__main__":
    main()
