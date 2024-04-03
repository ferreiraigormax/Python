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
def organizar():
    print('Você escolheu organizar, aguarde...')
    # Definindo as planilhas e diretórios
    planilha_veiculos = pd.read_excel(r'\Manutenção\Ferramentas\Cadastros\Cadastros.ods', engine='odf', sheet_name = 'Cadastro de Veículos')
    planilha_operadores = pd.read_excel(r'\Manutenção\Ferramentas\Cadastros\Cadastros.ods', engine='odf', sheet_name = 'Cadastro de Funcionários')
    diretorio = r'\Manutenção\Arquivo Digital\Relatórios'
    
    # Adicionando informações as listas
    for placas in planilha_veiculos['Placa']:
        lista_placas.append(placas)
    for operadores in planilha_operadores['Código']:
        lista_operadores.append(str(operadores))
    
    arquivos = os.listdir(f'{diretorio}\Organizar') # Definindo os arquivos que serão organizados
    for arquivo in arquivos:
        try:
            nome_arquivo_bruto = arquivo.split('.')[0]
            nome_arquivo = nome_arquivo_bruto.split(';')
            if len(nome_arquivo) == 6:
                placas_operadores_arquivo = nome_arquivo[1]
                if placas_operadores_arquivo in lista_placas or placas_operadores_arquivo in lista_operadores or 'CAP' in placas_operadores_arquivo:
                    mes_arquivo = nome_arquivo[3]
                    ano_arquivo = nome_arquivo[5]
                    if mes_arquivo in lista_mes and ano_arquivo in lista_ano:
                        shutil.move(f'{diretorio}\Organizar\{arquivo}', f'{diretorio}\{ano_arquivo}\{mes_arquivo}\{arquivo}')
                    else:
                        print(f'Erro com o mes ou ano --> {arquivo}')
                else:
                    print(f'Erro com a placa ou código do operador --> {arquivo}')
                    lista_arquivos_incorretos.append(arquivo)
            else:
                print(f'Erro com o nome do arquivo, ponto e vírgula --> {arquivo}')
                lista_arquivos_incorretos.append(arquivo)
        except:
            pass
# Menu principal
while True:
    print(''' 
Organizador de Relatórios Diários v1.0 | Escolha uma opção inserindo um dos seguintes números abaixo:''')
    
    # Definindo as listas
    lista_placas = []
    pool_placas = []
    lista_operadores = []
    lista_arquivos_incorretos = []
    lista_arquivos_corretos = []
    lista_mes = ['01', '02', '03', '04', '05', '06', '07', '08', '09','10', '11', '12']
    lista_ano = ['2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030'] #Criei uma lista, pois queria ter o controle dos anos em questão
    
    # Opções de menu principal
    menu = int(input('1 - Organizar\n2 - Sair\nDigite -> '))
    if menu == 1:
        organizar()
    else:
        print('Você escolheu sair...')
        break
