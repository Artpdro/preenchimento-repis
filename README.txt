Gerador de REPIS - Sindnorte PE
===============================

Aplicacao desktop (Python/tkinter) que preenche automaticamente o PDF modelo
de Certificado de Adesao ao REPIS: Nome da empresa, CNPJ e Codigo REPIS
(gerado automaticamente).

Estrutura
---------
main.py              -> ponto de entrada
interface.py         -> interface grafica
pdf_handler.py       -> preenchimento do PDF (overlay, original preservado)
repis_generator.py   -> geracao do Codigo REPIS
validators.py        -> validacao/mascara de CNPJ
config.py            -> coordenadas dos campos (ajuste fino aqui)
template/            -> PDF modelo
output/              -> PDFs gerados
tools/calibrate.py   -> previa PNG para conferir o alinhamento dos campos

Uso
---
1. pip install -r requirements.txt
2. (opcional, recomendado) python tools/calibrate.py  -> confira output/_preview.png
3. python main.py

Dependencias do sistema (Linux): python3-tk
  Debian/Ubuntu: sudo apt install python3-tk
