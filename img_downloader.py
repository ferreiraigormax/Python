from requests import get
from tkinter import *
def baixar(url):
    with open(r'pato.jpg', 'wb') as imagem:
      resposta = requests.get(url, stream=True)

      if not resposta.ok:
        print("Ocorreu um erro, status:" , resposta.status_code)
      else:
        for dado in resposta.iter_content(1024):
          if not dado:
              break

          imagem.write(dado)

        print("Imagem salva! =)")

# Criando o GUI para o usuário
root = Tk()
root.title('Baixar imagem v1.0') # Titulo da janela
root.geometry('260x30+0+0') # Tamanho da janela

# Definindo os Widgets
label1 = Label(root, text='Insira o Link ->')
entry1 = Entry(root)
button1 = Button(root, text='Baixar')

# Organizando os Widgets
label1.grid(row=1, column=0)
entry1.grid(row=1, column=1)
button1.grid(row=1, column=2)

root.mainloop() # Necessário para rodar a janela em loop
