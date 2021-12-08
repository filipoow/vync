import pyautogui
import time
import os

#Importação Tabelas
from config import dict_package
#Condições
numIdentificador = 0

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
        print("Origem {}".format(selectTable))
        print("\nPor gentileza solicitar testes.")
        print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
    #Email de uma tabela com taxa
    elif len(nomeTabela) == 1 and taxa > 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foi cadastrada no CNPJ {} a tabela:".format(cnpj))
        print("{} + {}".format(selectType,taxa))
        print("Origem {}".format(selectTable))
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
        print("Origem {}".format(selectTable))
        print("\nPor gentileza solicitar testes.")
    elif len(nomeTabela) == 1 and taxa == 0:
        print("\n-------------------------------------")
        print("            Script Email             ")
        print("-------------------------------------")
        print("{}"", boa tarde!".format(nomeVendedor))
        print("\nConforme solicitado foi cadastrada no CNPJ {} a tabela:".format(cnpj))
        print("{}".format(selectType))
        print("Origem {}".format(selectTable))
        print("\nPor gentileza solicitar testes.")
    #Caso de erro
    else:
        print("Sem informação para compilar o email, por favor entrar em contato com o criador.")
        


print("------------------------------------------------")
print("             Bem-vindo(a) ao Vync               ")
print("------------------------------------------------")
print("1 - Cadastro")
print("2 - Taxa")
opcao = int(input("Digite a operação(Ex: 1, 2): "))
clear_console()
if opcao == 1:
    #Requisitos para o loop...
    sair = "N"
    while sair == "N":
        #Se for colocar os códigos manualmente, adicionar antes de compilar..
        cod = []
        nomeTabela = {}

        print("------------------------------------------------")
        print("Preencher informações abaixo para o AutoCadastro")
        print("------------------------------------------------")
        nomeVendedor = input("Informar nome do vendedor subordinado: ")
        cnpj = input("Informe o cnpj: ")
        opt_code = 'S'
        while opt_code == 'S':
            #Selecionando o tipo da tabela
            key = "key{}".format(len(cod))
            selectType = input("Digite o tipo da tabela: ")
            nomeTabela[key] = selectType
            #Adicionando o código selecionado -- Package Balcão
            if selectType.lower() == 'package balcão':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['PackageBalcao']:
                    tableItem = dict_package['PackageBalcao'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 1
            if selectType.lower() == 'package 1':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package1']:
                    tableItem = dict_package['Package1'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- .COM 1 
            elif selectType.lower() == '.com 1':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['.Com1']:
                    tableItem = dict_package['.Com1'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 1,5 
            elif selectType.lower() == 'package 1.5':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package1.5']:
                    tableItem = dict_package['Package1.5'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- .COM 1,5
            elif selectType.lower() == '.com 1.5':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['.Com1.5']:
                    tableItem = dict_package['.Com1.5'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 2 
            elif selectType.lower() == 'package 2':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package2']:
                    tableItem = dict_package['Package2'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 2 Fronteira
            elif selectType.lower() == 'package 2 fronteira':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package2Front']:
                    tableItem = dict_package['Package2Front'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- .COM 2
            elif selectType.lower() == '.com 2':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['.Com2']:
                    tableItem = dict_package['.Com2'][selectTable]
                    cod.append(tableItem)
                elif selectTable == 'Multi':
                    listaCom2 = dict_package['.Com2'].values()
                    listaCom2 = list(listaCom2)
                    for item in listaCom2:
                        cod.append(item)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- .COM 2 Fronteira
            elif selectType.lower() == '.com 2 fronteira':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['.Com2Front']:
                    tableItem = dict_package['.Com2Front'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 4
            elif selectType.lower() == 'package 4':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package4']:
                    tableItem = dict_package['Package4'][selectTable]
                    cod.append(tableItem)
                elif selectTable == 'Multi':
                    listaPackage4 = dict_package['Package4'].values()
                    listaPackage4 = list(listaPackage4)
                    for item in listaPackage4:
                        cod.append(item)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- Package 4 Fronteira
            elif selectType.lower() == 'package 4 fronteira':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Package4Front']:
                    tableItem = dict_package['Package4Front'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            #Adicionando o código selecionado -- .COM 2
            elif selectType.lower() == 'pickup':
                selectTable = input("Digite a origem da tabela: ")
                if selectTable in dict_package['Pickup']:
                    tableItem = dict_package['Pickup'][selectTable]
                    cod.append(tableItem)
                else:
                    print("Erro na procura do código, informar o criador.")
            opt_code = input("Deseja cadastrar mais tabelas?(S/N) ")
        #Listagem dos nomes das tabelas...
        listaChave = nomeTabela.values()
        listaChave = list(listaChave)

        main()
        clear_console()
        print("Tabela cadastrada com sucesso!")
        print("------------------------------------------------")
        tableTaxa = input("Deseja adicionar taxa?(S/N) ")
        print("------------------------------------------------")
        if tableTaxa == 'S':
            #Entrada de dados...
            print("Tipo de Taxa: PACKAGE, .COM, PICKUP, RODOVIARIO, ECONOMICO, \nEXPRESSO, CORPORATE, DOC, CARGO, EMERGENCIAL")
            tipoTaxa = input("Digite o tipo da taxa: ")
            taxa = input("Digite a taxa(Ex: 0.75, 1.95, ...): ")
            addTaxa()
            taxa = float(taxa)
            clear_console()
            print("Taxa cadastrada com sucesso!!")
        else:
            clear_console()
            taxa = 0
            print('\nTaxa não adicionada.')
        scriptEmail()
        print("------------------------------------------------")
        sair = input("Deseja sair do programa?(S/N) ")
        print("------------------------------------------------")
        clear_console()
if opcao == 2:
    sair = "N"
    while sair == "N":
        print("----------------------------------------------------")
        print("Preencher informações abaixo para o cadastro de taxa")
        print("----------------------------------------------------")
        cnpj = input("Informe o cnpj: ")
        #Entrada de dados...
        print("Tipo de Taxa: PACKAGE, .COM, PICKUP, RODOVIARIO, ECONOMICO, \nEXPRESSO, CORPORATE, DOC, CARGO, EMERGENCIAL")
        tipoTaxa = input("Digite o tipo da taxa: ")
        taxa = input("Digite a taxa(Ex: 0.75, 1.95, ...): ")
        addTaxa()
        taxa = float(taxa)
        print("Taxa cadastrada com sucesso!!")
        print("------------------------------------------------")
        sair = input("Deseja sair do programa?(S/N) ")
        print("------------------------------------------------")
        clear_console()