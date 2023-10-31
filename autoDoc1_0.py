import tkinter
import PyPDF2
from docx import Document
import tkinter as tk  # Importe tkinter aqui
from tkinter import Tk, TkVersion, filedialog
from tkinter import simpledialog
from tkinter import messagebox
import os
import re
from docx2pdf import convert
from PIL import Image, ImageTk

# Função para preencher campos em um arquivo PDF
def preencher_pdf(arquivo_pdf):
    pdf_reader = PyPDF2.PdfFileReader(arquivo_pdf)
    pdf_writer = PyPDF2.PdfFileWriter()

    campos_pdf = {}

    # Detectar campos entre {{}}
    for page_num in range(pdf_reader.numPages):
        page = pdf_reader.getPage(page_num)
        page_text = page.extract_text()

        matches = re.finditer(r'{{(.*?)}}', page_text)
        for match in matches:
            campo = simpledialog.askstring('Editar campo', f'Editar campo: {match.group(1)}', parent=window)
            campos_pdf[match.group()] = campo

        page = pdf_reader.getPage(page_num)  # Obter a página original
        pdf_page = PyPDF2.PdfFileWriter()
        pdf_page.addPage(page)

        for campo, valor in campos_pdf.items():
            page_text = page_text.replace(campo, valor)
            pdf_page.getPage(0).mergePage(page)

        # Adicionar a página com os campos preenchidos ao novo PDF
        pdf_writer.addPage(pdf_page.getPage(0))

    # Gerar um novo arquivo PDF com o sufixo "Editado"
    output_pdf = os.path.splitext(arquivo_pdf)[0] + 'Editado.pdf'
    with open(output_pdf, 'wb') as output:
        pdf_writer.write(output)
    messagebox.showinfo("PDF Salvo", "O arquivo PDF foi preenchido e salvo com sucesso!")

# Função para preencher campos em um arquivo do Word
def preencher_word(arquivo_docx):
    doc = Document(arquivo_docx)

    campos_docx = {}

    # Detectar campos entre {{}}
    for para in doc.paragraphs:
        text = para.text
        matches = re.finditer(r'{{(.*?)}}', text)
        for match in matches:
            campo = simpledialog.askstring('Editar campo', f'Editar campo: {match.group(1)}', parent=window)
            campos_docx[match.group()] = campo

        for campo, valor in campos_docx.items():
            text = text.replace(campo, valor)

        para.clear()  # Limpar o texto original
        para.text = text

    # Gerar um novo arquivo do Word com o sufixo "Editado"
    output_docx = os.path.splitext(arquivo_docx)[0] + 'Editado.docx'
    doc.save(output_docx)
    messagebox.showinfo("Documento do Word Salvo", "O arquivo do Word foi preenchido e salvo com sucesso!")

# Função para converter um arquivo do Word para PDF
def converter_word_para_pdf():
    arquivo_docx = filedialog.askopenfilename(title="Selecione um arquivo do Word", filetypes=[("Documentos do Word", "*.docx")])
    if arquivo_docx:
        try:
            convert(arquivo_docx)
            messagebox.showinfo("Documento do Word Convertido para PDF", "O arquivo do Word foi convertido para PDF e salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro na conversão", f"Ocorreu um erro ao converter o arquivo: {str(e)}")

# Função para selecionar um arquivo
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo")
    if arquivo.lower().endswith('.pdf'):
        preencher_pdf(arquivo)
    elif arquivo.lower().endswith('.docx'):
        preencher_word(arquivo)
    else:
        messagebox.showinfo("Formato não suportado", "Formato de arquivo não suportado.")

# Função para exibir um guia passo a passo
def exibir_guia():
    guia = """
    --------------------------------------------
    Preparação do arquivo:
    1º - Antes de anexar o arquivo, substitua os campos de 
    preenchimento por Chaves {{ }}
    --------------------------------------------
    Passo 1: Abra o programa.
    Passo 2: Clique no botão "Selecionar Arquivo" para escolher
    o arquivo que deseja preencher.
    Passo 3: Siga as instruções na janela para preencher o 
    arquivo.
    Passo 4: Ao final do preenchimento o arquivo será salvo
    automaticamente em formado WORD(.docx).
    -------------------------------------------
    Bônus: Caso queira, é possível converter qualquer arquivo
    word(.docx) em PDF, basta clicar no botão:

    --------->>>"Converter WORD para PDF"<<<------------
    """
    messagebox.showinfo("Guia Passo a Passo", guia)
# Criar janela principal
window = tk.Tk()
window.title("Preenchimento de Documentos")

# Definir as dimensões da janela
largura_janela = 800
altura_janela = 600

# Obter as dimensões da tela
largura_tela = window.winfo_screenwidth()
altura_tela = window.winfo_screenheight()

# Calcular as coordenadas para centralizar a janela
x_pos = (largura_tela - largura_janela) // 2
y_pos = (altura_tela - altura_janela) // 2

# Configurar a janela com as dimensões desejadas
window.geometry(f"{largura_janela}x{altura_janela}+{x_pos}+{y_pos}")

# Definir um estilo de fonte
font_style = ("Helvetica", 12)  # Pode ajustar o tamanho da fonte aqui

# Converter a imagem JPEG para GIF
jpeg_image = Image.open("rocket.jpg")  # Substitua "rocket.jpg" pelo nome da sua imagem JPEG
gif_image = ImageTk.PhotoImage(jpeg_image)

# Configurar uma imagem de fundo
background_label = tk.Label(window, image=gif_image)
background_label.place(relwidth=1, relheight=1)

# Botão para selecionar arquivo
selecionar_arquivo_button = tk.Button(window, text="Selecionar Arquivo", command=selecionar_arquivo, font=font_style)
selecionar_arquivo_button.pack()

# Botão para converter Word para PDF
converter_word_para_pdf_button = tk.Button(window, text="Converter Word para PDF", command=converter_word_para_pdf, font=font_style)
converter_word_para_pdf_button.pack()

# Botão para exibir o guia passo a passo
exibir_guia_button = tk.Button(window, text="Exibir Guia Passo a Passo", command=exibir_guia, font=font_style)
exibir_guia_button.pack()

# Iniciar a janela
window.mainloop()