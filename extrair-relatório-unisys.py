'''Automação para extração de relatórios do Unisys
    Criado por: Igor Max - Manutenção do Transporte
    Contato: (16) 3946-4200 ou (16) 99138-1508'''

# Importando as bibliotecas
from pyautogui import press, write, click,dragTo, doubleClick, alert, hotkey, locateCenterOnScreen
from time import sleep
from pandas import read_excel
import clipboard as cl

# Definindo as datas
datas = ['0101202231122022']

# Defininfo os equipamentos
planilha = read_excel(r'C:\Users\igormax\Desktop\Manutenção\Ferramentas\Cadastros\Cadastros.ods', engine='odf') # Carrega os dados da planilha

# Funções responsáveis por extrair os relatórios ===========================================================================
def REPFT006F(): # 1 - REPFT006F - Requisição de veículos por veículo
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT006F.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                click(x, y) # Clica no push button Consumo de Combustível
                sleep(1)
                press('tab')
                write('EMUNA24') # Imprimi para o e-mail
                press('tab')
                press('tab')
                write(ano) # Datas
                write('2')
                press('tab')
                write(str(frota)) # Lista de Veículos
                press('tab')
                press('enter')
                sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT010A(): # 2 - REPFT010A - Relação notas fiscais por veículos
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT010A.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write('0')
                    press('tab')
                    write(ano) # Datas
                    write('0')
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('tab')
                    press('tab')
                    press('tab')
                    write('S')
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT010D(): # 3 - REPFT010D - Resumo analitico de despesas e receitas
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT010D.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write('1')
                    press('tab')
                    write(ano) # Datas
                    write('2')
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('tab')
                    press('tab')
                    press('tab')
                    write('S')
                    press('tab')
                    write('S')
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT010E(): # 4 - REPFT010E - Consumo de combustível
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT010E.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write('1')
                    press('tab')
                    write(ano) # Datas
                    write('2')
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('tab')
                    press('tab')
                    press('tab')
                    write('GERAL')
                    press('tab')
                    write('2')
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT010F(): # 5 - REPFT010F - Resumo analítico de despesas
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT010F.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write('1')
                    press('tab')
                    write(ano) # Datas
                    write('2')
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('enter')
                    sleep(8)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT013A(): # 6 - REPFT013A - Receitas por veículos
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT013A.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('tab')
                    write(ano) # Datas
                    write('0')
                    press('tab')
                    write('0')
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT022B(): # 7 - REPFT022B - Despesas mensais por veículos
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT022B.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    press('tab')
                    write(ano) # Datas
                    write('2')
                    press('tab')
                    write(str(frota)) # Lista de Veículos
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT010C(): # 8 - REPFT010C - Notas fiscais bloqueadas
    for a in range(0, 1):
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT010C.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Encontrando o botão para clicar
                    sleep(1)    
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    press('tab')
                    write(ano)
                    write('2')
                    press('enter')
                    sleep(6)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def REPFT013B(): # 9 - REPFT013B - Comissão por motoristas
    for frota in planilha['Frota']: # Conta quantos veículos existem no arquivo Frota.txt
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\REPFT013B.png')
            for ano in datas: # Lista de anos de 2015 ~ 2020
                    click(x, y) # Clica no push button Consumo de Combustível
                    sleep(1)
                    press('tab')
                    write('EMUNA24') # Imprimi para o e-mail
                    press('tab')
                    write('0')
                    press('tab')
                    write(ano) # Datas
                    press('tab')
                    press('tab')
                    press('tab')
                    press('tab')
                    press('tab')
                    write('G')
                    press('enter')
                    sleep(10)
        except:
            alert(text='Não foi possível encontrar o botão !', title='Atenção!', button='Ok')
            break
    alert(text='Extração dos relatórios do Unisys realizada com sucesso...', title='Fim', button='Ok')
def salvar(ano, equipamento, relatorio): # Salvando os arquivos
    x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\top_bar.png')
    click(x - 50, y)
    sleep(2)
    caminho = (fr'C:\Users\igormax\Desktop\Backup Relatórios do Unisys\{ano}\{equipamento}\{relatorio}')
    sleep(1)
    cl.copy(caminho)
    sleep(1)
    hotkey('ctrl', 'v')
    press('enter')
    sleep(1)
    x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\botom_bar.png')
    click(x, y)
    sleep(2)
    nome = (fr'{equipamento} - {relatorio} ({ano})')
    cl.copy(nome)
    sleep(1)
    hotkey('ctrl', 'v')
    sleep(1)
    press('enter')
    sleep(3)
    try:
        x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\alert.png')
        press('left')
        press('enter')
        sleep(0.5)
        hotkey('alt', 'f4')
        hotkey('alt', 'f4')
        sleep(0.5)
    except:
        hotkey('alt', 'f4')
        hotkey('alt', 'f4')
        sleep(1)
def nao_salvar(): # Não salvar o arquivo, casos em que o relatório sai em branco
    hotkey('alt', 'f4')
    hotkey('alt', 'f4')
    sleep(0.5)
def emails(): # Funções responsáveis por organizar os e-mails
    while True:
        press('down')
        press('enter')
        #try:
        #    x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\encaminhar.png')
        #    click(x, y)
        #    sleep(1)
        #except:
        #    alert('Não foi possível encaminhar o e-mail!')
        #    break
        sleep(1)
        try:
            x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\anexo.png')
            click(x, y)
        except:
            alert('Não foi possível encontrar o anexo!')
            break
        press('enter')
        sleep(1)
        press('up')
        sleep(1)
        press('enter')
        sleep(1)
        x, y = locateCenterOnScreen(r'C:\Users\igormax\Documents\Projetos\Python\Unisys_backup\img\max.png')
        click(x, y)
        click(0, 50)
        dragTo(0, 300, 1)
        sleep(0.5)
        hotkey('ctrl', 'c')
        teste = cl.paste()
        try:
            if 'REPFT006F' in teste:
                click(742, 109)
                sleep(1)
                dragTo(778, 109, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(118, 222)
                sleep(1)
                dragTo(149, 222, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT006F - Requisição de Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT010A' in teste:
                click(658, 108)
                sleep(1)
                dragTo(696, 108, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(21, 261)
                sleep(1)
                dragTo(49, 261, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT010A - Relação de Notas Fiscais por Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT010D' in teste:
                click(732, 280)
                sleep(1)
                dragTo(769, 280, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(568, 402)
                sleep(1)
                dragTo(598, 402, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT010D - Relação Analítica de Despesas e Receitas por Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT010E' in teste:
                click(885, 116)
                sleep(1)
                dragTo(920, 116, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(202, 213)
                sleep(1)
                dragTo(229, 213, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT010E - Relação de Consumo de Combustível' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT010F' in teste:
                click(806, 116)
                sleep(1)
                dragTo(839, 116, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(82, 211)
                sleep(1)
                dragTo(111, 211, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT010F - Resumo Analítico de Despesas por Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT013A' in teste:
                click(634, 97)
                sleep(1)
                dragTo(672, 97, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(102, 184)
                sleep(1)
                dragTo(130, 184, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT013A - Receita de Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT022B' in teste:
                click(956, 116)
                sleep(1)
                dragTo(995, 116, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                click(400, 210)
                sleep(1)
                dragTo(428, 210, 1)
                hotkey('ctrl', 'c')
                equipamento = cl.paste()
                sleep(2)
                if equipamento == ano:
                    nao_salvar()
                else:
                    relatorio = 'REPFT022B - Despesa Mensal por Veículos' # Localização do equipamento
                    hotkey('ctrl','shift', 's')
                    sleep(2)
                    salvar(ano, equipamento, relatorio)
            elif 'REPFT010C' in teste: # Relatório geral, não tem frota em especifico 
                click(742, 109)
                sleep(1)
                dragTo(778, 109, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                equipamento = '000 - Geral'
                sleep(2)
                relatorio = 'REPFT010C - Relação de Notas Fiscais Bloqueadas' # Localização do equipamento
                hotkey('ctrl','shift', 's')
                sleep(2)
                salvar(ano, equipamento, relatorio)
            elif 'REPFT013B' in teste: # Relatório geral, não tem frota em especifico 
                click(659, 98)
                sleep(1)
                dragTo(700, 98, 1)
                hotkey('ctrl', 'c')
                ano = cl.paste()
                sleep(2)
                equipamento = '000 - Geral'
                sleep(2)
                relatorio = 'REPFT013B - Relação de Comissão por Motoristas' # Localização do equipamento
                hotkey('ctrl','shift', 's')
                sleep(2)
                salvar(ano, equipamento, relatorio)               
            else:
                alert('O processo falhou!')
                break
        except:
            alert('error')
while True:
    print('''-----------SISTEMA DE IMPRESSÃO AUTOMATIZADA UNISYS-----------
                  :!!!!!!!!!!!!!!!!!!!!!!!                   
              :!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!:              
           !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!           
         !!                      .!!!!!                     
       :!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!:       
      !!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!      
     :!!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!!     
     !!!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!!     
     !!!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!!     
     :!!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!:     
      !!!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!!!      
       :!!!!!!!!            .!!!!!!!!!!!!!!!    !!!!:       
         !!!!!!!            .!!!!!!!!!!!!!!!    !!!         
           !!!!!!            !!!!!!!!!!!!!!:   :!           
              :!!!            .!!!!!!!!!!.                  
                   :!.                       :..:  ''')
    try:
        menu = int(input('''-------------Escolha uma das opções para imprimir-------------
        
    1 - REPFT006F - Requisição de veículos por veículo
    2 - REPFT010A - Relação notas fiscais por veículos
    3 - REPFT010D - Resumo analitico de despesas e receitas
    4 - REPFT010E - Consumo de combustível
    5 - REPFT010F - Resumo analítico de despesas
    6 - REPFT013A - Receitas por veículos
    7 - REPFT022B - Despesas mensais por veículos
    8 - REPFT010C - Notas fiscais bloqueadas
    9 - REPFT013B - Comissão por motoristas
    10 - Organizar os e-mails
    0 - Sair
    >> '''))
        if menu == 1:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT006F()
        elif menu == 2:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT010A()
        elif menu == 3:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT010D()
        elif menu == 4:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT010E()
        elif menu == 5:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT010F()
        elif menu == 6:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT013A()
        elif menu == 7:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT022B()
        elif menu == 8:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT010C()
        elif menu == 9:
            hotkey('alt', 'tab')
            sleep(1)
            REPFT013B()
        elif menu == 10:
            hotkey('alt', 'tab')
            sleep(1)
            emails()
        elif menu == 0:
            print('Você escolheu sair, saindo...')
            break
    except:
            alert(text='Opção inválida, tente novamente!', title='Atenção!', button='Ok')
            pass


