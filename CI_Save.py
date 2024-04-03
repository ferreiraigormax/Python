'''Criando um programa que tira um print da tela e depois salva.'''
# Importando as bibliotecas necessárias
from pyautogui import screenshot
from datetime import datetime
from time import sleep
from tkinter import *
import os

# Definindo a função do button1
def capturar():
    captura = screenshot()
    hoje = datetime.today().strftime('%d-%m-%Y')
    try:
        os.mkdir(fr'C:\Users\igormax\Desktop\CI\{hoje}')
    except:
        pass
    sleep(1)
    captura.save(fr'C:\Users\igormax\Desktop\CI\{hoje}\{hoje} {entry1.get()}.jpeg')
    pressionado.set(str('Sucesso!'))
    local = (fr'C:\Users\igormax\Desktop\CI\{hoje}')
    local = os.path.realpath(local)
    os.startfile(local)

# Definindo as configurações da tela
root = Tk()
root.title(r'Salvar C.I (Copercana) v1.0')
root.geometry('310x50+0+0')

# Definindo as Strings Variveis
pressionado = StringVar()

# Criando os Widgets
label1 = Label(root, text='Nome do Arquivo: ', font='Arial, 8')
label2 = Label(root, textvariable=pressionado)
label3 = Label(root, text='Status ->', font='Arial, 8')
entry1 = Entry(root)
button1 = Button(root, text='PrintScreen', command=capturar)

# Organizando os Widgets
label1.grid(row=0, column=0)
label2.grid(row=1, column=1)
label3.grid(row=1, column=0)
entry1.grid(row=0, column=1)
button1.grid(row=0, column=2)

# Usado para rodar constantemente
root.mainloop()
