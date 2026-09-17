"""Gera uma previa PNG do modelo com textos de teste nas posicoes de config.py.

Execute:  python tools/calibrate.py
Abra  output/_preview.png  e confira o alinhamento. Se necessario, ajuste
FIELDS em config.py e rode novamente.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from PIL import Image, ImageDraw
from PyPDF2 import PdfReader

from config import FIELDS, TEMPLATE_PDF

SAMPLE = {
    "nomeemp": "EMPRESA EXEMPLO LTDA",
    "cnpj": "12.345.678/0001-90",
    "codrepis": "RP-AB12CD34",
}


def main():
    page = PdfReader(str(TEMPLATE_PDF)).pages[0]

    # Extrai a imagem da pagina
    xo = page["/Resources"]["/XObject"].get_object()
    form = next(iter(xo.values())).get_object()
    inner = next(iter(form["/Resources"]["/XObject"].values())).get_object()
    img = Image.frombytes("RGB", (int(inner["/Width"]), int(inner["/Height"])),
                          inner.get_data())

    w_pt, h_pt = float(page.mediabox.width), float(page.mediabox.height)
    sx, sy = img.width / w_pt, img.height / h_pt

    draw = ImageDraw.Draw(img)
    for field, text in SAMPLE.items():
        cfg = FIELDS[field]
        x = cfg["x"] * sx
        y_img = img.height - cfg["y"] * sy - cfg["size"] * sy
        draw.text((x, y_img), text, fill=(0, 0, 0))

    out = Path(__file__).resolve().parent.parent / "output" / "_preview.png"
    out.parent.mkdir(exist_ok=True)
    img.save(out)
    print(f"Previa salva em: {out}")


if __name__ == "__main__":
    main()
