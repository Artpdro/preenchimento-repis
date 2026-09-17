"""Importacao de empresas a partir de planilha Excel (.xlsx).

Funcoes separadas: leitura/preview da planilha e importacao para o banco,
com deduplicacao por CNPJ e relatorio de novos/atualizados/ignorados.
"""

import re
import unicodedata
from pathlib import Path

import openpyxl

import database
from validators import only_digits

# Mapeamento flexivel de cabecalhos (sem acento, minusculo) -> campo interno
_HEADER_MAP = {
    "cnpj": "cnpj",
    "cnpj da empresa": "cnpj",
    "nome": "nome",
    "nome da empresa": "nome",
    "empresa": "nome",
    "razao social": "nome",
    "razao_social": "nome",
    "codigo": "codigo_repis",
    "codigo repis": "codigo_repis",
    "codigo_repis": "codigo_repis",
    "repis": "codigo_repis",
    "registro": "codigo_repis",
    "n registro": "codigo_repis",
    "no registro": "codigo_repis",
    "numero": "codigo_repis",
}


def _norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text.strip().lower())


def read_preview(path: str | Path, max_linhas: int = 50) -> dict:
    """
    Le o arquivo e retorna dict com:
      colunas   -> lista de colunas encontradas no cabecalho
      mapeamento-> dict campo_interno -> indice da coluna
      linhas    -> lista de dicts (amostra, ate max_linhas)
      erros     -> lista de mensagens de problema estrutural
    """
    resultado = {"colunas": [], "mapeamento": {}, "linhas": [], "erros": []}
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        resultado["erros"].append(f"Nao foi possivel abrir o arquivo: {exc}")
        return resultado

    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        resultado["erros"].append("A planilha esta vazia.")
        wb.close()
        return resultado

    colunas = [str(h).strip() if h is not None else "" for h in header]
    resultado["colunas"] = colunas
    mapeamento = {}
    for idx, col in enumerate(colunas):
        campo = _HEADER_MAP.get(_norm(col))
        if campo and campo not in mapeamento:
            mapeamento[campo] = idx
    resultado["mapeamento"] = mapeamento

    if "cnpj" not in mapeamento or "nome" not in mapeamento:
        resultado["erros"].append(
            "Colunas obrigatorias nao encontradas. A planilha deve possuir "
            "cabecalhos para CNPJ e Nome da empresa (ou Razao Social). "
            f"Cabecalhos encontrados: {', '.join(colunas) or '(nenhuno)'}")
        wb.close()
        return resultado

    amostra = []
    for i, row in enumerate(rows):
        if i >= max_linhas:
            break
        reg = {"_linha": i + 2}
        for campo, idx in mapeamento.items():
            valor = row[idx] if idx < len(row) else None
            reg[campo] = "" if valor is None else str(valor).strip()
        amostra.append(reg)
    wb.close()
    resultado["linhas"] = amostra
    return resultado


def importar(path: str | Path, progresso=None) -> dict:
    """
    Importa o arquivo para o banco. Retorna relatorio:
      novos, atualizados, ignorados (lista de motivos), total_processado
    """
    relatorio = {"novos": 0, "atualizados": 0, "ignorados": [], "total": 0}

    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        relatorio["ignorados"].append(f"Arquivo invalido: {exc}")
        return relatorio

    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        wb.close()
        relatorio["ignorados"].append("A planilha esta vazia.")
        return relatorio

    colunas = [str(h).strip() if h is not None else "" for h in header]
    mapeamento = {}
    for idx, col in enumerate(colunas):
        campo = _HEADER_MAP.get(_norm(col))
        if campo and campo not in mapeamento:
            mapeamento[campo] = idx

    if "cnpj" not in mapeamento or "nome" not in mapeamento:
        wb.close()
        relatorio["ignorados"].append(
            "Importacao cancelada: colunas obrigatorias (CNPJ e Nome) ausentes.")
        return relatorio

    for i, row in enumerate(rows):
        relatorio["total"] += 1
        nome = str(row[mapeamento["nome"]] or "").strip()
        cnpj = only_digits(str(row[mapeamento["cnpj"]] or ""))
        codigo = ""
        if "codigo_repis" in mapeamento:
            codigo = str(row[mapeamento["codigo_repis"]] or "").strip()

        if not nome and not cnpj:
            relatorio["ignorados"].append(f"Linha {i + 2}: linha vazia.")
            continue
        if len(cnpj) != 14:
            relatorio["ignorados"].append(
                f"Linha {i + 2}: CNPJ invalido ({cnpj or 'vazio'}).")
            continue
        try:
            status, _ = database.upsert_empresa(nome, cnpj, codigo)
            relatorio["novos" if status == "novo" else "atualizados"] += 1
        except ValueError as exc:
            relatorio["ignorados"].append(f"Linha {i + 2}: {exc}")
        if progresso:
            progresso(i + 1)
    wb.close()
    return relatorio
