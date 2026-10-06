"""Ponto de entrada executável da automação.

Uso:
    python main.py --help
    python main.py sync --mes 7 --ano 2026
    python main.py sync --teste
    python main.py resetar-teste
    python main.py concluir "Nome" "CODIGO" --data 07/08/2026
"""

import os
import sys
from pathlib import Path


# Permite executar este arquivo diretamente, sem instalação prévia do pacote.
RAIZ_ATUAL = Path(__file__).resolve().parent
if str(RAIZ_ATUAL) not in sys.path:
    sys.path.insert(0, str(RAIZ_ATUAL))


# Modo "--teste": redireciona a base pro sandbox _Automação/_ambiente-teste/
# (cópias das planilhas reais). Isso PRECISA acontecer antes de a cli importar o
# treinamentos_its.config, porque os caminhos do config são resolvidos no import.
#
# Quem decide "isto é modo teste" é `ambiente_teste.modo_teste_pedido`, e SÓ ele:
# a regra que estava escrita aqui na mão discordava do argparse (que aceitava
# `--tes`, `--test`, ...) e fazia um "modo teste" rodar contra as planilhas de
# PRODUÇÃO. O `cli` ainda confirma, antes de executar, que a base virou mesmo o
# sandbox — se estas duas pontas discordarem, ele aborta em vez de escrever.
from treinamentos_its.ambiente_teste import modo_teste_pedido, preparar_sandbox  # noqa: E402

if modo_teste_pedido(sys.argv):
    os.environ["TREINAMENTOS_ITS_BASE"] = str(preparar_sandbox())

from treinamentos_its.cli import main  # noqa: E402


if __name__ == "__main__":
    main()
