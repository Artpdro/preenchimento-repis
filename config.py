"""Configuracoes e coordenadas dos campos (unico ponto de ajuste).

MAPEAMENTO REAL - certificado_sem_dados.pdf
Pagina unica: 593.28 x 440.64 pt | Imagem 2472 x 1836 px | escala 0.24 pt/px
Barras detectadas por analise de pixels (preenchimento solido RGB ~190):
  nomeemp : x 492-2321 px, y 933-989 px
  cnpj    : x 960-1451 px, y 991-1043 px
  codrepis: barra apos "sob nº" -> x 328-1022 px, y 1147-1202 px  (ver print 1)
"""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PDF = BASE_DIR / "template" / "certificado_sem_dados.pdf"
OUTPUT_DIR = BASE_DIR / "output"
LOGO_PNG = BASE_DIR / "assets" / "logo.png"

PAGE_WIDTH = 593.28
PAGE_HEIGHT = 440.64

# Coordenadas em pontos PDF, origem no canto INFERIOR esquerdo.
# x = borda esquerda da barra (+4 pt de margem interna); y = baseline do texto.
FIELDS = {
    "nomeemp":  {"x": 122.0, "y": 206.0, "size": 10.0, "max_chars": 70},
    "cnpj":     {"x": 234.0, "y": 193.0, "size": 9.5,  "max_chars": 18},
    "codrepis": {"x": 82.0,  "y": 155.0, "size": 10.0, "max_chars": 14},
}

REPIS_PREFIX = "RP"
REPIS_LENGTH = 8  # caracteres alfanumericos apos o prefixo

# Janela principal
WINDOW_WIDTH = 1366
WINDOW_HEIGHT = 768
