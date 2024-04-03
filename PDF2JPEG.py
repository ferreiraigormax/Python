'''Criando um programa para converter PDF para JPEG'''

from tkinter import *
from tkinter.filedialog import askdirectory
import os
from pdf2image import convert_from_path


def converter():
    path = askdirectory()
    files = os.listdir(path)
    for file in files:
        print(file)
        img = convert_from_path(fr'{path}/{file}', dpi=200)

root = Tk()
root.title('PDF2JPEG vers.1.0')

descricao = Label(root, text="Escolha uma pasta para converter os arquivos .pdf para .jpeg")
botao = Button(root, text="Abrir", command=converter)

descricao.pack()
botao.pack()
root.mainloop()

# Pegar os arquivos

# Converter os arquivos

# Excluir arquivos anteriores

# Abrir pasta do arquivo
