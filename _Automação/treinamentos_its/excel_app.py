"""Abertura padronizada do Excel via xlwings.

Todo mundo que abre uma planilha deve passar por aqui. Com o Excel invisível, um
pop-up de "atualizar vínculos?" ou um alerta qualquer trava o processo sem
ninguém ver (o `sync` fica pendurado pra sempre). A gente desliga de uma vez:

- ``update_links=False`` na abertura -> não pergunta sobre vínculos externos;
- ``display_alerts=False`` -> sem caixinha de confirmação;
- ``screen_updating=False`` -> não tenta redesenhar a tela (app invisível).

O preço de desligar os alertas: se o arquivo estiver aberto em outro computador
(ou em outro Excel desta máquina), o Excel abre SOMENTE LEITURA **em silêncio**
— sem o aviso "arquivo em uso". Quem ia escrever não percebia nada. Por isso o
``abrir_livro`` confere o ``ReadOnly`` de verdade quando a abertura é pra
escrita e recusa na hora (``ArquivoAbertoSomenteLeitura``).

O ``import xlwings`` fica dentro das funções de propósito, pra este módulo
continuar importável (e o pacote continuar testável) numa máquina sem Excel.
O ``reprocessamento.py`` já fazia o certo na mão; agora é tudo por aqui.
"""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path


class ArquivoAbertoSomenteLeitura(RuntimeError):
    """O workbook foi pedido pra escrita, mas o Excel abriu somente leitura."""


@contextmanager
def app_excel(visible: bool = False):
    """Abre uma instância do Excel já 'calada' e garante o fechamento no fim."""
    import xlwings as xw

    app = xw.App(visible=visible)
    try:
        # display_alerts/screen_updating não são desligados sozinhos pelo xlwings.
        app.display_alerts = False
        app.screen_updating = False
        yield app
    finally:
        _encerrar_excel(app)


# Quantas vezes o xlwings repete uma chamada que o Excel recusou por estar
# "ocupado" ENQUANTO a gente tenta fechá-lo. O padrão do xlwings é 0 = infinito,
# sem pausa: um Excel preso num diálogo fazia o `quit()` girar pra sempre e o
# `kill()` (e a restauração depois dele) nunca chegava (revisão do Codex,
# 05/10/2026). Só no FECHAR: no resto (recálculo, colagem) o Excel fica ocupado
# alguns segundos de forma legítima e um limite geral derrubaria a operação.
TENTATIVAS_COM_AO_FECHAR = 50


@contextmanager
def _tentativas_com_limitadas(limite: int):
    """Limita, só dentro do bloco, as repetições do xlwings em "Excel ocupado"."""
    try:
        import xlwings._xlwindows as xlwindows
    except Exception:  # noqa: BLE001 - sem o xlwings real (testes, outra plataforma)
        yield
        return
    antes = getattr(xlwindows, "N_COM_ATTEMPTS", 0)
    xlwindows.N_COM_ATTEMPTS = limite
    try:
        yield
    finally:
        xlwindows.N_COM_ATTEMPTS = antes


def _encerrar_excel(app) -> None:
    """Fecha o Excel e, se ele não responder, mata o processo.

    O ``with xw.App()`` do xlwings só chega no ``kill()`` se o ``quit()`` der
    certo. Em 05/10/2026 o Excel invisível travou num diálogo do OneDrive que o
    ``display_alerts=False`` não segura ("couldn't find '...sharepoint.../...
    (2).xlsx'" — conflito de nome na sincronização): o ``quit()`` estourou
    ``0x800AC472`` (Excel ocupado) e o processo ficou vivo, invisível, preso no
    diálogo. Aqui o ``kill()`` roda sempre, e um erro ao fechar nunca esconde o
    erro que interrompeu a operação.
    """
    from treinamentos_its.logger import log

    quit_falhou = False
    try:
        with _tentativas_com_limitadas(TENTATIVAS_COM_AO_FECHAR):
            app.quit()
    except Exception as erro:  # noqa: BLE001
        quit_falhou = True
        log.warning(
            "O Excel não fechou normalmente (%s); encerrando o processo dele.", erro
        )
    try:
        app.kill()
    except Exception as erro:  # noqa: BLE001
        # Depois de um quit() bom o processo já saiu e o kill() falha à toa. Depois
        # de um quit() que FALHOU, não: o Excel pode ter ficado vivo segurando a
        # planilha (revisão do Codex, 05/10/2026) — isso não pode passar calado.
        if quit_falhou:
            pid = getattr(getattr(app, "impl", None), "_pid", None)
            log.error(
                "NÃO consegui encerrar o Excel invisível (processo %s): %s. Feche-o "
                "pelo Gerenciador de Tarefas (EXCEL.EXE sem janela) antes de rodar de "
                "novo.",
                pid if pid is not None else "?",
                erro,
            )


def _fechar_sem_salvar(wb) -> None:
    """Fecha o workbook recusado; um erro aqui não pode esconder o motivo real."""
    try:
        wb.close()
    except Exception:  # noqa: BLE001
        pass


def abrir_livro(app, caminho, *, read_only: bool = False):
    """Abre um workbook sem disparar o pop-up de vínculos externos.

    Use sempre isto no lugar de ``app.books.open(...)`` cru.

    Com ``read_only=False`` (quem vai ESCREVER), confere se o Excel abriu mesmo
    pra escrita. Abriu somente leitura (arquivo aberto em outro lugar, arquivo
    marcado como somente leitura, OneDrive segurando) -> fecha e levanta
    ``ArquivoAbertoSomenteLeitura``. Se nem der pra ler o ``ReadOnly``, também
    recusa: sem essa confirmação a escrita poderia "dar certo" sem gravar nada.
    """
    wb = app.books.open(caminho, update_links=False, read_only=read_only)
    if read_only:
        return wb

    nome = Path(str(caminho)).name
    try:
        somente_leitura = bool(wb.api.ReadOnly)
    except Exception as erro:  # noqa: BLE001
        _fechar_sem_salvar(wb)
        raise ArquivoAbertoSomenteLeitura(
            f"Não consegui confirmar que '{nome}' abriu para ESCRITA ({erro}). Nada "
            f"foi alterado nele. Feche o Excel e rode de novo."
        ) from erro
    if somente_leitura:
        _fechar_sem_salvar(wb)
        raise ArquivoAbertoSomenteLeitura(
            f"'{nome}' abriu SOMENTE LEITURA — está aberto em outro computador/Excel? "
            f"Feche e rode de novo. Nada foi alterado nele."
        )
    return wb
