'''
	Titulo: Organizador de Notas Fiscais
	Criado por: Igor Max Ferreira
	Ult.modificação: 03/08/2022
'''

# Importando as bibliotecas
import os
import pandas as pd
import shutil

# Definindo funções
def mover(dir, arq, mes, ano):
    try:
        shutil.move(fr'{dir}\Organizar\{arq}', fr'{dir}\{ano}\{mes}\{arq}')
    except:
        print(f'Falha ao mover o {arq}')

def organizar():
    print('Você escolheu organizar, aguarde...')
    # Definindo as planilhas e diretórios
    planilha_veiculos = pd.read_excel(r'\Manutenção\Ferramentas\Cadastros\Cadastros.ods', engine='odf', sheet_name = 'Cadastro de Veículos')
    planilha_fornecedores = pd.read_excel(r'Manutenção\Ferramentas\Cadastros\Cadastros.ods', engine='odf', sheet_name = 'Cadastro de Fornecedores')
    diretorio = r'Manutenção\Arquivo Digital\Notas Fiscais\Consumo'
    
    # Adicionando informações as listas
    for placas in planilha_veiculos['Placa']:
        lista_placas.append(placas)
    for fornecedores in planilha_fornecedores['Código']:
        lista_fornecedores.append(str(fornecedores))
    
    arquivos = os.listdir(f'{diretorio}\organizar') # Definindo os arquivos que serão organizados
    for arquivo in arquivos:
        try:
            nome_arquivo_bruto = arquivo.split('.')[0]
            nome_arquivo = nome_arquivo_bruto.split(';')
            if len(nome_arquivo) > 1 and len(nome_arquivo) < 13:
                placas_arquivo = nome_arquivo[11]
                fornecedor_arquivo = nome_arquivo[9]
                data_arquivo = nome_arquivo[1]
                try:
                    mes = data_arquivo.split('-')[1]
                    ano = data_arquivo.split('-')[2]
                except:
                    pass
                if len(data_arquivo) != 10: #Separar o que estiver com a data errada!
                    lista_arquivos_incorretos.append(f'Erro Data: {data_arquivo} --> {arquivo}')
                elif fornecedor_arquivo not in lista_fornecedores: #Separar o que estiver com o fornecedor errado!
                    lista_arquivos_incorretos.append(f'Erro Fornecedor: {fornecedor_arquivo} --> {arquivo}')
                elif placas_arquivo not in lista_placas: #Separar o que estiver com a placa errada!
                    try:
                        contagem = 0
                        for a in placas_arquivo.split(','):
                            if a not in lista_placas:
                                lista_arquivos_incorretos.append(f'Erro Placa: {a} --> {arquivo}')
                            else:
                                contagem += 1
                        if contagem == len(placas_arquivo.split(',')):
                            mover(diretorio, arquivo, mes, ano)
                            lista_arquivos_corretos.append(arquivo)
                        else:
                            pass
                    except:
                        lista_arquivos_incorretos.append(f'Erro Placa: {placas_arquivo} --> {arquivo}')
                else:
                    mover(diretorio, arquivo, mes, ano)
                    lista_arquivos_corretos.append(arquivo)
            else:
                lista_arquivos_incorretos.append(f'Erro com: ";" --> {arquivo}')
        except:
            pass
            
            

    '''print('\nArquivos corretos: ')
    for arquivo1 in lista_arquivos_corretos:
        print(arquivo1)'''
        
    print('\nArquivos incorretos: ')
    for arquivo2 in lista_arquivos_incorretos:
        print(arquivo2)
# Menu principal
while True:
    print('''   
Organizador de Notas Fiscais v1.0 | Escolha uma opção inserindo um dos seguintes números abaixo:''')
    
    # Definindo as listas
    lista_placas = ['-', 'DIVERSAS']
    lista_fornecedores = ['COPERCANA']
    lista_arquivos_incorretos = []
    lista_arquivos_corretos = []
    
    # Opções de menu principal
    menu = int(input('1 - Organizar\n2 - Sair\nDigite -> '))
    if menu == 1:
        organizar()
    else:
        print('Você escolheu sair...')
