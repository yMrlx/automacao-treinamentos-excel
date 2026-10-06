"""Transações físicas para arquivos Excel sujeitos ao AutoSave.

O ciclo mensal e o ``concluir`` alteram somente um arquivo de Planos e Macs.
Fechar o workbook sem ``save()`` não é rollback nesta máquina: o OneDrive pode
persistir uma escrita intermediária. Por isso a garantia continua sendo uma
cópia física restaurada somente depois que o Excel fecha o arquivo.
"""

from __future__ import annotations

import shutil
from contextlib import ExitStack
from pathlib import Path

from treinamentos_its.logger import log


CODIGO_SAIDA_RESTAURACAO_INCOMPLETA = 3


class ErroAoPrepararRestauracao(RuntimeError):
    """Não foi possível criar a garantia física antes de alterar um arquivo."""


def _caminho_disponivel(destino: Path) -> Path:
    if not destino.exists():
        return destino
    contador = 2
    while True:
        candidato = destino.with_name(f"{destino.stem} ({contador}){destino.suffix}")
        if not candidato.exists():
            return candidato
        contador += 1


def copiar_para_restauracao(
    arquivo: str | Path, destino: str | Path, *, descricao: str
) -> Path | None:
    arquivo = Path(arquivo)
    destino = _caminho_disponivel(Path(destino))
    if not arquivo.exists():
        log.error("Arquivo a proteger não encontrado: %s", arquivo.resolve())
        return None
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(arquivo, destino)
    except Exception as erro:  # noqa: BLE001
        log.exception(
            "Não consegui criar a cópia de restauração de %s em '%s': %s",
            descricao,
            destino,
            erro,
        )
        return None
    log.info("Cópia de restauração de %s: %s", descricao, destino.resolve())
    return destino


def restaurar_arquivo(
    copia: str | Path, arquivo: str | Path, *, descricao: str
) -> bool:
    copia = Path(copia)
    arquivo = Path(arquivo)
    try:
        shutil.copy2(copia, arquivo)
    except Exception as erro:  # noqa: BLE001
        log.exception(
            "RESTAURAÇÃO FALHOU para %s ('%s'): %s. NÃO rode o ciclo novamente "
            "antes de restaurar manualmente a partir de '%s'.",
            descricao,
            arquivo,
            erro,
            copia,
        )
        return False
    log.warning(
        "%s foi RESTAURADO a partir de '%s'; nenhuma alteração desta operação foi mantida.",
        descricao,
        copia,
    )
    return True


def arquivos_em_uso(caminhos) -> list[Path]:
    """Arquivos que outro programa (o Excel de alguém) está segurando agora.

    Abrir para leitura+escrita sem escrever nada não altera o arquivo; se o
    Windows recusa, alguém está com ele aberto (ou ele está travado/somente
    leitura). Medido em 05/10/2026: com o Setembro aberto no Excel, o `concluir`
    não conseguia gravar, tentava restaurar a cópia por cima do arquivo travado,
    falhava e terminava com RESTAURAÇÃO INCOMPLETA (código 3) — com o arquivo
    intacto. Checar ANTES evita mexer e dá a mensagem certa.
    """
    em_uso: list[Path] = []
    for caminho in caminhos:
        caminho = Path(caminho)
        if not caminho.exists():
            continue
        try:
            with open(caminho, "r+b"):
                pass
        except PermissionError:
            em_uso.append(caminho)
    return em_uso


def avisar_arquivos_em_uso(em_uso: list[Path]) -> None:
    for caminho in em_uso:
        log.error(
            "'%s' está ABERTO no Excel (seu ou de outra pessoa) ou travado pelo "
            "OneDrive. Feche esse arquivo em todos os computadores e rode de novo. "
            "Nada foi alterado.",
            caminho.name,
        )


def desligar_autosave(wb, *, descricao: str) -> bool:
    """Tenta desligar o AutoSave; a cópia física segue sendo a garantia."""
    try:
        if getattr(wb.api, "AutoSaveOn", False):
            wb.api.AutoSaveOn = False
            log.info("AutoSave do OneDrive DESLIGADO em %s durante a operação.", descricao)
        return True
    except Exception as erro:  # noqa: BLE001
        log.warning(
            "Não consegui desligar o AutoSave em %s (%s). A operação segue "
            "protegida pela cópia física de restauração.",
            descricao,
            erro,
        )
        return False


class TransacaoArquivo:
    """Restaura o arquivo se a operação não for confirmada.

    Se o destino não existia ao entrar, rollback significa apagar o arquivo
    incompleto criado pela operação.
    """

    def __init__(
        self,
        arquivo: str | Path,
        destino_copia: str | Path,
        *,
        descricao: str,
        exigir_existente: bool = True,
    ) -> None:
        self.arquivo = Path(arquivo)
        self.destino_copia = Path(destino_copia)
        self.descricao = descricao
        self.exigir_existente = exigir_existente
        self.existia = False
        self.copia: Path | None = None
        self.restauracao_falhou = False
        self._confirmada = False

    def __enter__(self):
        self.existia = self.arquivo.exists()
        if self.existia:
            self.copia = copiar_para_restauracao(
                self.arquivo, self.destino_copia, descricao=self.descricao
            )
            if self.copia is None:
                raise ErroAoPrepararRestauracao(
                    f"Sem cópia de restauração, não vou alterar {self.descricao}."
                )
        elif self.exigir_existente:
            raise ErroAoPrepararRestauracao(
                f"Arquivo de {self.descricao} não encontrado: {self.arquivo.resolve()}"
            )
        return self

    def confirmar(self) -> None:
        self._confirmada = True

    @property
    def confirmada(self) -> bool:
        return self._confirmada

    @property
    def falhas_de_restauracao(self) -> list["TransacaoArquivo"]:
        return [self] if self.restauracao_falhou else []

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        if self._confirmada:
            return False
        if self.existia and self.copia is not None:
            if not restaurar_arquivo(self.copia, self.arquivo, descricao=self.descricao):
                self.restauracao_falhou = True
        elif self.arquivo.exists():
            try:
                self.arquivo.unlink()
                log.warning(
                    "Arquivo incompleto de %s removido após falha: %s",
                    self.descricao,
                    self.arquivo,
                )
            except Exception as erro:  # noqa: BLE001
                self.restauracao_falhou = True
                log.exception(
                    "NÃO consegui remover o arquivo incompleto de %s ('%s'): %s. "
                    "Apague-o manualmente antes de rodar de novo.",
                    self.descricao,
                    self.arquivo,
                    erro,
                )
        return False


class TransacaoArquivos:
    """Compatibilidade para operações que protegem mais de um arquivo."""

    def __init__(self, *transacoes: TransacaoArquivo) -> None:
        self.transacoes = transacoes
        self._pilha: ExitStack | None = None

    def __enter__(self):
        pilha = ExitStack()
        try:
            for transacao in self.transacoes:
                pilha.enter_context(transacao)
        except BaseException:
            pilha.close()
            raise
        self._pilha = pilha
        return self

    def confirmar(self) -> None:
        for transacao in self.transacoes:
            transacao.confirmar()

    @property
    def confirmada(self) -> bool:
        return all(transacao.confirmada for transacao in self.transacoes)

    @property
    def falhas_de_restauracao(self) -> list[TransacaoArquivo]:
        return [
            falha
            for transacao in self.transacoes
            for falha in transacao.falhas_de_restauracao
        ]

    @property
    def restauracao_falhou(self) -> bool:
        return bool(self.falhas_de_restauracao)

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        if self._pilha is None:
            return False
        return self._pilha.__exit__(exc_type, exc_value, traceback)


def instrucoes_de_restauracao_manual(falhas: list[TransacaoArquivo]) -> list[str]:
    instrucoes: list[str] = []
    for transacao in falhas:
        if transacao.copia is not None:
            instrucoes.append(
                f"'{transacao.arquivo}': com o Excel FECHADO, copie '{transacao.copia}' "
                "por cima dele"
            )
        else:
            instrucoes.append(
                f"'{transacao.arquivo}': nasceu nesta execução e ficou pela metade - "
                "apague-o à mão"
            )
    return instrucoes
