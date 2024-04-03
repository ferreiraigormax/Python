'''
	Titulo: Transformar notas fiscais de .xml para .csv
	Criado por: Igor Max Ferreira
	Ult.modificação: 04/10/2022
'''
# Importando as bibliotecas necessárias
import csv
import xml.etree.ElementTree as ET
import os
  
arquivos = os.listdir(r'C:\Users\igormax\Documents\Projetos\Python\xml\xml2csv\Transformar') # <- Definir o local dos arquivos
for arquivo in arquivos:
    try:
        tree = ET.parse(fr'C:\Users\igormax\Documents\Projetos\Python\xml\xml2csv\Transformar\{arquivo}')
        root = tree.getroot()
        cnpj_fornecedor = root.findall(".//{http://www.portalfiscal.inf.br/nfe}CNPJ")[0].text
        cnpj_cliente = root.findall(".//{http://www.portalfiscal.inf.br/nfe}CNPJ")[1].text
        nota_fiscal = root.findall(".//{http://www.portalfiscal.inf.br/nfe}nNF")[0].text
        data_emissao = root.findall(".//{http://www.portalfiscal.inf.br/nfe}dhEmi")[0].text
        placa = root.findall(".//{http://www.portalfiscal.inf.br/nfe}placa")[0].text
        nome_items = root.findall(".//{http://www.portalfiscal.inf.br/nfe}xProd")
        qtd_items = root.findall(".//{http://www.portalfiscal.inf.br/nfe}qCom")
        valor_items = root.findall(".//{http://www.portalfiscal.inf.br/nfe}vUnCom")
        dia = data_emissao[8:10] #Definindo o dia da Nota Fiscal
        mes = data_emissao[5:7] #Definindo o mês da Nota Fiscal
        ano = data_emissao[0:4] #Definindo o ano da Nota Fiscal
        data = (f'{dia}/{mes}/{ano}') #Data combinada no padrão dd/mm/aaaa
        with open(r'C:\Users\igormax\Documents\Projetos\Python\xml\notas.csv', 'a', newline='') as csvfile:
            escrever = csv.writer(csvfile, delimiter=';')
            for c in range(0, len(nome_items)):
                try:
                    item = nome_items[c].text.replace('.', ',')
                    qtd = qtd_items[c].text.replace('.', ',')
                    valor = valor_items[c].text.replace('.', ',')
                    escrever.writerow(['cnpj_cliente', 'cnpj_fornecedor', 'nota_fiscal', data, 'item', qtd, valor, 'placa'])
                except:
                    print('Erro - Item não encontrado')
                    pass
    except:
        print('Erro - XML incorreto')
        pass
