#Importação de tabela por pandas
import pandas as pd
import pyautogui
import time
import os

#Importação Tabelas
from config import dict_package

df = pd.read_excel('base.xlsx', index_col=False)
# Convertendo df para dicionário
base = df.to_dict('index')
i = 0 

def clear_console():
    os.system('cls')

def addTaxa():
    # #Entrada de dados...
    # print("Tipo de Taxa: PACKAGE, .COM, PICKUP, RODOVIARIO, ECONOMICO, \nEXPRESSO, CORPORATE, DOC, CARGO, EMERGENCIAL")
    # tipoTaxa = input("Digite o tipo da taxa: ")
    # taxa = input("Digite a taxa: ")

    #Abrindo o Cadastro Central
    pyautogui.click(x=217, y=748)
    #Abrindo aba de taxa
    pyautogui.click(x=36, y=33)
    time.sleep(1)
    pyautogui.press('tab',presses=2)
    pyautogui.press('enter')
    time.sleep(1)
    #Inserindo CNPJ e localizando CNPJ
    pyautogui.click(x=522, y=283)
    pyautogui.write(cnpj)
    # pyautogui.press('tab')
    # pyautogui.press('enter')
    #Selecionando a tipo de taxa
    if tipoTaxa.lower() == "package":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando Package
        pyautogui.click(x=598, y=497)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)
    if tipoTaxa.lower() == ".com":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando .COM
        pyautogui.click(x=569, y=588)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)
    if tipoTaxa.lower() == "rodoviario":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando RODOVIARIO
        pyautogui.click(x=570, y=513)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)
    if tipoTaxa.lower() == "economico":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando ECONOMICO
        pyautogui.click(x=573, y=533)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)
    if tipoTaxa.lower() == "expresso":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando EXPRESSO
        pyautogui.click(x=568, y=475)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)
    if tipoTaxa.lower() == "pickup":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando PICKUP
        pyautogui.click(x=570, y=641)
        time.sleep(1)
        #Inserindo taxa
        pyautogui.click(x=561, y=492)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        pyautogui.typewrite(taxa)
        #Salvando taxa..
        pyautogui.click(x=855, y=456)
        #Prevenção de erros
        pyautogui.click(x=706, y=430)
        time.sleep(1)
        #Atualizando dados..
        pyautogui.click(x=715, y=279)


def main():
    #Abrindo o Cadastro Central
    pyautogui.click(x=217, y=748)
    #Loop para cadastro das tabelas..
    for c in range(len(cod)):
        pyautogui.click(x=403, y=362)
        #Prevenção de erro
        with pyautogui.hold('ctrl'):
            pyautogui.press('a')
        pyautogui.press('backspace')
        time.sleep(1)
        pyautogui.typewrite(cod[c])
        time.sleep(1)
        #Vincular essa tabela
        pyautogui.click(x=432, y=402)
        time.sleep(1)
        #Inserindo CNPJ
        pyautogui.click(x=424, y=341)
        pyautogui.write(cnpj)
        time.sleep(1)
        #Buscando CNPJ
        pyautogui.click(x=623, y=340)
        time.sleep(1)
        #Salvando o vínculo
        pyautogui.click(x=878, y=343)
        time.sleep(1)
        #Prevenção de erro
        pyautogui.click(x=760, y=440)
        time.sleep(1)
    #Atualizando 
    pyautogui.click(x=421, y=124)
    #Prevenção de erro
    with pyautogui.hold('ctrl'):
        pyautogui.press('a')
    pyautogui.press('backspace')
    time.sleep(1)
    pyautogui.write(cnpj)
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('enter')

def scriptEmail():
    #Entrada de dados..

    #Email de mais de uma tabela c/ taxa...
    if len(nomeTabela) > 1 and taxa > 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foram cadastradas no CNPJ {} as tabelas:".format(cnpj))
        for chave in listaChave:
            print("{} + {}".format(chave,taxa))
        print("Origem {}".format(origem))
        print("\nPor gentileza solicitar testes.")
        print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
    #Email de uma tabela com taxa
    elif len(nomeTabela) == 1 and taxa > 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foi cadastrada no CNPJ {} a tabela:".format(cnpj))
        print("{} + {}".format(tipoTabela,taxa))
        print("Origem {}".format(origem))
        print("\nPor gentileza solicitar testes.")
        print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
    #Email mais de uma tabela sem taxa.
    elif len(nomeTabela) > 1 and taxa == 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foram cadastradas no CNPJ {} as tabelas:".format(cnpj))
        for chave in listaChave:
            print("{}".format(chave))
        print("Origem {}".format(origem))
        print("\nPor gentileza solicitar testes.")
    elif len(nomeTabela) == 1 and taxa == 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foi cadastrada no CNPJ {} a tabela:".format(cnpj))
        print("{}".format(tipoTabela))
        print("Origem {}".format(origem))
        print("\nPor gentileza solicitar testes.")
    #Caso de erro
    else:
        print("Sem informação para compilar o email, por favor entrar em contato com o criador.")
        
while i < len(base):
    nomeVendedor = str(base[i]['nomeVendedor'])
    cnpj = str(base[i]['cnpj'])
    tipoTabela = str(base[i]['tipoTabela'])
    origem = str(base[i]['Origem'])
    seTaxa = str(base[i]['seTaxa'])
    tipoTaxa = str(base[i]['tipoTaxa'])
    taxa = str(base[i]['taxa'])
    
    #Se for colocar os códigos manualmente, adicionar antes de compilar..
    cod = []
    nomeTabela = {}
    #Selecionando o tipo da tabela
    key = "key{}".format(len(cod))
    nomeTabela[key] = tipoTabela
    #Adicionando o código selecionado -- Package Balcão
    if tipoTabela.lower() == 'package balcão':
        if origem in dict_package['PackageBalcao']:
            itemTabela = dict_package['PackageBalcao'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 1
    if tipoTabela.lower() == 'package 1':
        if origem in dict_package['Package1']:
            itemTabela = dict_package['Package1'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- .COM 1 
    elif tipoTabela.lower() == '.com 1':
        if origem in dict_package['.Com1']:
            itemTabela = dict_package['.Com1'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 1,5 
    elif tipoTabela.lower() == 'package 1.5':
        if origem in dict_package['Package1.5']:
            itemTabela = dict_package['Package1.5'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- .COM 1,5
    elif tipoTabela.lower() == '.com 1.5':
        if origem in dict_package['.Com1.5']:
            itemTabela = dict_package['.Com1.5'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 2 
    elif tipoTabela.lower() == 'package 2':
        if origem in dict_package['Package2']:
            itemTabela = dict_package['Package2'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 2 Fronteira
    elif tipoTabela.lower() == 'package 2 fronteira':
        if origem in dict_package['Package2Front']:
            itemTabela = dict_package['Package2Front'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- .COM 2
    elif tipoTabela.lower() == '.com 2':
        if origem in dict_package['.Com2']:
            itemTabela = dict_package['.Com2'][origem]
            cod.append(itemTabela)
        elif origem == 'Multi':
            listaCom2 = dict_package['.Com2'].values()
            listaCom2 = list(listaCom2)
            for item in listaCom2:
                cod.append(item)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- .COM 2 Fronteira
    elif tipoTabela.lower() == '.com 2 fronteira':
        if origem in dict_package['.Com2Front']:
            itemTabela = dict_package['.Com2Front'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 4
    elif tipoTabela.lower() == 'package 4':
        if origem in dict_package['Package4']:
            itemTabela = dict_package['Package4'][origem]
            cod.append(itemTabela)
        elif origem == 'Multi':
            listaPackage4 = dict_package['Package4'].values()
            listaPackage4 = list(listaPackage4)
            for item in listaPackage4:
                cod.append(item)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- Package 4 Fronteira
    elif tipoTabela.lower() == 'package 4 fronteira':
        if origem in dict_package['Package4Front']:
            itemTabela = dict_package['Package4Front'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Adicionando o código selecionado -- .COM 2
    elif tipoTabela.lower() == 'pickup':
        if origem in dict_package['Pickup']:
            itemTabela = dict_package['Pickup'][origem]
            cod.append(itemTabela)
        else:
            print("Erro na procura do código, informar o criador.")
    #Listagem dos nomes das tabelas...
    listaChave = nomeTabela.values()
    listaChave = list(listaChave)

    main()
    # clear_console()
    print("Tabela cadastrada com sucesso!")
    print("------------------------------------------------")
    if seTaxa.lower() == 's':
        addTaxa()
        taxa = float(taxa)
        # clear_console()
        print("Taxa cadastrada com sucesso!!")
    elif seTaxa.lower() == 'n':
        # clear_console()
        taxa = 0
        print('\nTaxa não adicionada.')
    else:
        print("Erro ao efetuar a ação, entrar em contato com o criador.")
    scriptEmail()
    print("------------------------------------------------")
    print("Todas tabelas cadastradas!")
    i += 1

