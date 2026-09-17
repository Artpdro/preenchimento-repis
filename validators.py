"""Validacao de CNPJ (com digitos verificadores, sem dependencias externas)."""

import re


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def format_cnpj(cnpj: str) -> str:
    """Aplica a mascara 00.000.000/0000-00."""
    digits = only_digits(cnpj)
    if len(digits) != 14:
        return digits
    return (f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/"
            f"{digits[8:12]}-{digits[12:14]}")


def is_valid_cnpj(cnpj: str) -> bool:
    digits = only_digits(cnpj)
    if len(digits) != 14 or len(set(digits)) == 1:
        return False

    def dv(nums, weights):
        s = sum(n * w for n, w in zip(nums, weights))
        r = s % 11
        return 0 if r < 2 else 11 - r

    d1 = dv([int(d) for d in digits[:12]], list(range(5, 1, -1)) + list(range(9, 1, -1)))
    d2 = dv([int(d) for d in digits[:13]], list(range(6, 1, -1)) + list(range(9, 1, -1)))
    return digits[-2:] == f"{d1}{d2}"


def validate_empresa(nome: str) -> str | None:
    nome = (nome or "").strip()
    if not nome:
        return "Informe o nome da empresa."
    if len(nome) > 120:
        return "Nome da empresa muito longo (max. 120 caracteres)."
    return None
