"""Historico de certificados emitidos (persistido em output/historico.json)."""

import json
from datetime import datetime

from config import OUTPUT_DIR

HISTORY_FILE = OUTPUT_DIR / "historico.json"


def load_history() -> list[dict]:
    """Retorna a lista de registros (mais recente primeiro)."""
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def add_record(empresa: str, cnpj: str, codigo: str, arquivo) -> dict:
    """Adiciona um registro ao historico e retorna o registro criado."""
    record = {
        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "empresa": empresa,
        "cnpj": cnpj,
        "codigo": codigo,
        "arquivo": str(arquivo),
    }
    historico = load_history()
    historico.insert(0, record)
    HISTORY_FILE.write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")
    return record


def clear_history() -> None:
    if HISTORY_FILE.exists():
        HISTORY_FILE.unlink()
