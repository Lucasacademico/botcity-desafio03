from PyPDF2 import PdfWriter
from PyPDF2 import PdfReader
from PyPDF2 import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

def milimetro_pontos(milimetro):
    return milimetro / 0.352777

def criar_arquivo():
    # Criando um objeto PdfWriter
    pdf_writer = PdfWriter()
    # Tamanho A4 em pontos (1 ponto = 1/72 polegada)
    nova_pagina = PageObject.create_blank_page(width=595.2, height=841.8)
    # Nova_pagina
    # Adicionando uma página ao arquivo PDF
    pdf_writer.add_page(nova_pagina)
    # Salvando o arquivo PDF no disco
    with open("meu_arquivo.pdf", "wb") as output_pdf:
        pdf_writer.write(output_pdf)

def criar_pdf():
    largura, altura = A4
    pdf = canvas.Canvas('./teste.pdf', pagesize=A4)
    #pdf.setFont('Helvetica-Oblique', 16)
    pdf.setFont('Helvetica', 16)
    titulo = 'Avaliacao parte 01'
    print(len(titulo))
    pdf.drawString( ((largura / 2) - (len(titulo) / 2)), altura - 90, titulo)
    pdf.save()