"""Camada de acesso ao banco de dados (SQLite, sem dependencias externas).

Toda a persistencia de empresas fica neste modulo. Futuras melhorias
(novos campos, outros modelos de PDF) devem estender as funcoes daqui.
"""

import re
import sqlite3
from datetime import datetime
from pathlib import Path

from config import OUTPUT_DIR

DB_PATH = OUTPUT_DIR / "empresas.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS empresas (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    nome          TEXT NOT NULL,
    cnpj          TEXT NOT NULL UNIQUE,
    codigo_repis  TEXT,
    criado_em     TEXT NOT NULL,
    atualizado_em TEXT NOT NULL
);
"""


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute(_SCHEMA)
    return conn


def _now() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


# ---------------------------------------------------------------- CRUD
def upsert_empresa(nome: str, cnpj: str, codigo_repis: str = "") -> tuple[str, dict]:
    """
    Insere ou atualiza uma empresa (CNPJ e o identificador principal).
    Retorna ("novo"|"atualizado", registro). Levanta ValueError se dados invalidos.
    """
    nome = (nome or "").strip()
    cnpj_digits = only_digits(cnpj)
    if not nome:
        raise ValueError("Nome da empresa ausente.")
    if len(cnpj_digits) != 14:
        raise ValueError(f"CNPJ invalido: {cnpj!r}")

    cnpj_fmt = (f"{cnpj_digits[0:2]}.{cnpj_digits[2:5]}.{cnpj_digits[5:8]}/"
                f"{cnpj_digits[8:12]}-{cnpj_digits[12:14]}")
    agora = _now()
    with _connect() as conn:
        row = conn.execute("SELECT * FROM empresas WHERE cnpj = ?",
                           (cnpj_fmt,)).fetchone()
        if row:
            novo_codigo = codigo_repis or row["codigo_repis"] or ""
            conn.execute(
                "UPDATE empresas SET nome = ?, codigo_repis = ?, atualizado_em = ? "
                "WHERE cnpj = ?", (nome, novo_codigo, agora, cnpj_fmt))
            status = "atualizado"
        else:
            conn.execute(
                "INSERT INTO empresas (nome, cnpj, codigo_repis, criado_em, atualizado_em)"
                " VALUES (?, ?, ?, ?, ?)",
                (nome, cnpj_fmt, codigo_repis or "", agora, agora))
            status = "novo"
        registro = conn.execute("SELECT * FROM empresas WHERE cnpj = ?",
                                (cnpj_fmt,)).fetchone()
    return status, dict(registro)


def get_empresa_by_cnpj(cnpj: str) -> dict | None:
    cnpj_fmt = only_digits(cnpj)
    if len(cnpj_fmt) != 14:
        return None
    cnpj_fmt = (f"{cnpj_fmt[0:2]}.{cnpj_fmt[2:5]}.{cnpj_fmt[5:8]}/"
                f"{cnpj_fmt[8:12]}-{cnpj_fmt[12:14]}")
    with _connect() as conn:
        row = conn.execute("SELECT * FROM empresas WHERE cnpj = ?",
                           (cnpj_fmt,)).fetchone()
    return dict(row) if row else None


def update_codigo_repis(cnpj: str, codigo: str) -> None:
    with _connect() as conn:
        conn.execute("UPDATE empresas SET codigo_repis = ?, atualizado_em = ? "
                     "WHERE cnpj = ?", (codigo, _now(), cnpj))


def search_empresas(termo: str, limite: int = 100) -> list[dict]:
    """Busca parcial por CNPJ (digitos) ou por codigo REPIS/nome."""
    termo = (termo or "").strip()
    digits = only_digits(termo)
    with _connect() as conn:
        if digits and len(digits) >= 3:
            like = "%" + "%".join(digits) + "%"
            rows = conn.execute(
                "SELECT * FROM empresas WHERE REPLACE(REPLACE(REPLACE(cnpj,'.',''),'/',''),'-','') "
                "LIKE ? ORDER BY nome LIMIT ?", (like, limite)).fetchall()
        else:
            like = f"%{termo}%"
            rows = conn.execute(
                "SELECT * FROM empresas WHERE nome LIKE ? OR codigo_repis LIKE ? "
                "ORDER BY nome LIMIT ?", (like, like, limite)).fetchall()
    return [dict(r) for r in rows]


def list_empresas(limite: int = 500) -> list[dict]:
    with _connect() as conn:
        rows = conn.execute("SELECT * FROM empresas ORDER BY nome LIMIT ?",
                            (limite,)).fetchall()
    return [dict(r) for r in rows]


def count_empresas() -> int:
    with _connect() as conn:
        return conn.execute("SELECT COUNT(*) c FROM empresas").fetchone()["c"]
