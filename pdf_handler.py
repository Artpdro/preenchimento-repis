"""Leitura do PDF modelo e preenchimento via overlay transparente.

O PDF original NUNCA e alterado: criamos uma camada separada com os textos
e a mesclamos sobre a pagina original.
"""

from datetime import datetime
from io import BytesIO
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.colors import black
from reportlab.pdfgen import canvas

from config import FIELDS, OUTPUT_DIR, PAGE_HEIGHT, PAGE_WIDTH, TEMPLATE_PDF


def fill_pdf(nome_empresa: str, cnpj: str, codigo_repis: str,
             registro_mte: str = "", output_path: Path | None = None) -> Path:
    if output_path is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "".join(c if c.isalnum() else "_" for c in nome_empresa)[:40]
        output_path = OUTPUT_DIR / f"REPIS_{safe_name}_{stamp}.pdf"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(TEMPLATE_PDF))
    page = reader.pages[0]

    # ---- Overlay transparente com apenas os textos dos campos ----
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    c.setFillColor(black)

    values = {
        "nomeemp": nome_empresa,
        "cnpj": cnpj,
        "codrepis": codigo_repis,
    }
    if registro_mte:
        values["registroMTE"] = registro_mte  # campo opcional (se existir em FIELDS)

    for field, text in values.items():
        if field not in FIELDS or not text:
            continue
        cfg = FIELDS[field]
        text = str(text)[: cfg["max_chars"]]  # respeita area disponivel
        c.setFont("Helvetica-Bold", cfg["size"])
        c.drawString(cfg["x"], cfg["y"], text)

    c.save()
    buf.seek(0)

    # ---- Mescla overlay sobre o original (original intacto) ----
    overlay = PdfReader(buf).pages[0]
    page.merge_page(overlay)

    writer = PdfWriter()
    writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

    return output_path
