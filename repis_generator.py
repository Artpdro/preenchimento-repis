"""Geracao do Codigo do REPIS - regra isolada para facil manutencao."""

import secrets
import string

from config import REPIS_LENGTH, REPIS_PREFIX

# Evita codigos duplicados dentro da mesma execucao.
_issued: set[str] = set()

_ALPHABET = string.ascii_uppercase + string.digits


def generate_repis_code() -> str:
    """
    Gera um codigo unico no formato RP-XXXXXXXX.

    Para alterar a regra (ex.: sequencial com banco de dados, ano embutido),
    basta modificar esta funcao - o restante do sistema nao muda.
    """
    while True:
        code = f"{REPIS_PREFIX}-{''.join(secrets.choice(_ALPHABET) for _ in range(REPIS_LENGTH))}"
        if code not in _issued:
            _issued.add(code)
            return code
