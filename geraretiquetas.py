"""Gera PDF com etiquetas QR a partir do CSV de itens.
Uso:  python gerar_etiquetas.py itens_exemplo.csv
Instalar:  pip install qrcode[pil] reportlab
"""
import csv, sys, io
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

arquivo = sys.argv[1] if len(sys.argv) > 1 else "itens_exemplo.csv"
COLS, LINHAS = 3, 8                      # etiquetas por página
LARG, ALT = A4[0] / COLS, A4[1] / LINHAS

with open(arquivo, encoding="utf-8-sig", newline="") as f:
    amostra = f.read(2048); f.seek(0)
    sep = ";" if ";" in amostra.splitlines()[0] else ","
    itens = list(csv.DictReader(f, delimiter=sep))

c = canvas.Canvas("etiquetas.pdf", pagesize=A4)
for n, it in enumerate(itens):
    pos = n % (COLS * LINHAS)
    if n and pos == 0:
        c.showPage()
    col, lin = pos % COLS, pos // COLS
    x, y = col * LARG, A4[1] - (lin + 1) * ALT
    img = qrcode.make(it["codigo"].strip(), box_size=10, border=1)
    buf = io.BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
    lado = min(ALT - 14 * mm, LARG * 0.45)
    c.drawImage(ImageReader(buf), x + 4 * mm, y + (ALT - lado) / 2, lado, lado)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x + lado + 6 * mm, y + ALT / 2 + 4 * mm, it["codigo"])
    c.setFont("Helvetica", 7)
    c.drawString(x + lado + 6 * mm, y + ALT / 2 - 3 * mm, (it.get("descricao") or "")[:28])
    c.rect(x + 1 * mm, y + 1 * mm, LARG - 2 * mm, ALT - 2 * mm)   # guia de corte
c.save()
print(f"{len(itens)} etiquetas em etiquetas.pdf")