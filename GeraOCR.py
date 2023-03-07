# BIBLIOTECAS #
import fitz
import os
import pytesseract
import cv2 
from tkinter import messagebox
import comtypes.client
# BIBLIOTECAS #

msg_cnv = messagebox.showinfo("Convertendo Imagem...", "Deixa comigo!")

# PARTE 1 AQUI É ONDE O OS E O FITZ REALIZARÁ SEU PRIMEIRO SHOW; #
    # Aqui o "Fitz" e o "OS" são responsáveis por gerenciar e converter os PDFs em Imagens; #
        # "OS" é a biblioteca que irá apontar o diretório dos PDF; #
            # Enquanto o "Fitz" é o cara que vai converter os PDFs. #

# DIRETÓRIO ONDE SE ENCONTRA O SEU PSD QUE SERÁ CONVERTIDO EM IMAGEM #            
pathpdfimg = R"C:\Users\Usuario\Desktop\1. AleRO\Projeto\ReconhecerImgNv" 
pdflist = os.listdir(pathpdfimg)
print(pdflist)

# !! CAMINHO DO TESSERACT PARA QUEM USA WINDOWS !! #
caminho = R"C:\Users\Usuario\AppData\Local\Programs\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + R"\Tesseract.exe"
# !! CAMINHO DO TESSERACT PARA QUEM USA WINDOWS !! #

for pdf in pdflist:
    if pdf.endswith('.pdf'): 
        doc = fitz.open(pdf) # "FITZ" ABRE O PDF #
        number_of_pages = doc.page_count # FAZ UMA CONTAGEM DE PÁGINAS #

        for pag in range(1,number_of_pages + 1): 
                    page = doc.load_page(pag -1) 
                    matriz = fitz.Matrix(8, 8)
                    pix = page.get_pixmap(matrix=matriz)
                    output = f'img_{pag}.jpg'
                    pix.save(output) # AQUI É ONDE A MÁGIA ACONTECE E ELE SAI COMO IMG #

# !! VÁRIAVEIS QUE CRIEI PRO PROGRAMA FUNCIONAR !! #
outpu = number_of_pages # DEI O MESMO VALOR PARA A VÁRIAVEL OUTPU PARA NÃO ALTERAR NADA NO ESCOPO #
pag_in_out = 0
outpi = 1
outpi1 = 1
pag_nam_out = 0
pag = (pag + 1) # ADICIONEI O VALOR +1 PARA FUNCIONAR NA VÁRIAVEL REMOVE DO "OS" (ELE APAGAVA SEMPRE O 'VALOR SAÍDA' OU O ÚLTIMO E DEPOIS O TESSERACT NÃO ENCONTRAVA O ARQUIVO) #
# !! VÁRIAVEIS QUE CRIEI PRO PROGRAMA FUNCIONAR !! #

doctexto = open("Template.txt", "a")

while (outpu > pag_in_out): # ENQUANTO OUTPU FOR MAIOR QUE PAG_IN_OUT ELE CONTINUARA EXECUTANDO O COMANDO ABAIXO #
    resul = pytesseract.image_to_string(f'img_{outpu}.jpg') # LÊ AS LINHAS DA IMAGEM(AQUI QUE A MAGIA ACONTE) #
    print(resul) # PRINTA O RESULTADO #
    doctexto.write(resul) # OUTPUT / BLOCO DE NOTAS #
    outpu = (outpu - 1) # DIMINUI O OUTPU QUE É RELACIONADO AO NÚMERO DE PÁGINAS (IMG_1...) #

while (outpi < pag): # ENQUANTO OUTPI FOR MENOR QUE PAG ELE EXECUTARA O COMANDO ABAIXO QUANTAS VEZES FOR PRECISO #
        os.remove(F'img_{outpi}.jpg') # LEMBRA QUE ELE CRIA IMAGENS NO DIRETÓRIO? #  # NESSA LINHA ELE APAGA TODAS AS IMAGENS CRIADAS PELO PROGRAMA # 
        outpi = (outpi + 1) # ADICIONA O VALOR PARA IR APAGANDO RELACIONADO AS PÁGINAS (IMG_1...) #  
    
else:
    path = pathpdfimg
    word = comtypes.client.CreateObject("Template.docx")
    word.Documents.Open(path,ReadOnly=1)
    word.Run("Project.Modulo1.NewMacro")
    word.Documents(1).Close(SaveChanges=0)
    word.Application.Quit()
    wd=0
    msg_fim = messagebox.showinfo("Atenção!", "Prontinho, Volte sempre!")    


