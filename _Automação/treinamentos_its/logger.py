"""Configuração única de logs da aplicação."""

import logging


def configurar_logging() -> None:
    """Configura o logger sem substituir handlers criados por um chamador."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S",
    )


configurar_logging()
log = logging.getLogger("treinamentos_its")
