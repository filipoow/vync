import pyautogui
import time
import os

#Importação Tabelas
from config import Package1
from config import Com1
from config import Com1_5
from config import Package2
from config import Package2Fronteira
from config import Com2
from config import Package4

#Condições
numIdentificador = 0

def clear_console():
    os.system('cls')

def addTaxa():
    #Entrada de dados...
    print("Tipo de Taxa: PACKAGE, .COM, PICKUP, RODOVIARIO, ECONOMICO, \nEXPRESSO, CORPORATE, DOC, CARGO, EMERGENCIAL")
    tipoTaxa = input("Digite o tipo da taxa: ")
    taxa = input("Digite a taxa: ")

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
    if tipoTaxa == "PACKAGE":
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
    if tipoTaxa == ".COM":
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
    if tipoTaxa == "RODOVIARIO":
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
    if tipoTaxa == "ECONOMICO":
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
    if tipoTaxa == "EXPRESSO":
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
    if tipoTaxa == "PICKUP":
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

def scriptEmail(selectType, selectTable, taxa):
    #Entrada de dados..
    print("\n-------------------------------------")
    print("            Script Email             ")
    print("-------------------------------------")
    # print("{}"", boa tarde!".format(nomeVendedor))
    print("\nConforme solicitado foram cadastradas no CNPJ {} as tabelas:".format(cnpj))
    print("{} + {}".format(selectType,taxa))
    print("Origem {}".format(selectTable))
    print("\nPor gentileza solicitar testes.")
    print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")

sair = "N"

while sair == "N":
    #Se for colocar os códigos manualmente, adicionar antes de compilar..
    cod = []

    print("------------------------------------------------")
    print("Preencher informações abaixo para o AutoCadastro")
    print("------------------------------------------------")
    # nomeVendedor = input("Informar nome do vendedor subordinado: ") Ainda testando....
    cnpj = input("Informe o cnpj: ")
    #Selecionando o tipo da tabela
    selectType = input("Digite o tipo da tabela: ")
    #Adicionando o código selecionado -- Package 1
    if selectType == 'Package 1':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Package1:
            PackageItem = Package1[selectTable]
            cod.append(PackageItem)
    #Adicionando o código selecionado -- .COM 1 
    elif selectType == '.COM 1':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Com1:
            ComItem = Com1[selectTable]
            cod.append(ComItem)
    #Adicionando o código selecionado -- .COM 1,5
    elif selectType == '.COM 1.5':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Com1_5:
            Com1_5Item = Com1_5[selectTable]
            cod.append(Com1_5Item)
    #Adicionando o código selecionado -- Package 2 
    elif selectType == 'Package 2':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Package2:
            Package2Item = Package2[selectTable]
            cod.append(Package2Item)
    #Adicionando o código selecionado -- Package 2 Fronteira
    elif selectType == 'Package 2 Fronteira':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Package2Fronteira:
            Package2FrontItem = Package2Fronteira[selectTable]
            cod.append(Package2FrontItem)
    #Adicionando o código selecionado -- .COM 2
    elif selectType == '.COM 2':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Com2:
            Com2Item = Com2[selectTable]
            cod.append(Com2Item)
    #Adicionando o código selecionado -- .COM 2
    elif selectType == 'Package 4':
        selectTable = input("Digite a origem da tabela: ")
        if selectTable in Package4:
            Package4Item = Package4[selectTable]
            cod.append(Package4Item)

    main()
    clear_console()
    print("Tabela cadastrada com sucesso!")
    time.sleep(1)
    clear_console()
    print("------------------------------------------------")
    tableTaxa = input("Deseja adicionar taxa?(S/N) ")
    print("------------------------------------------------")
    if tableTaxa == 'S':
        addTaxa()
        print("Taxa cadastrada com sucesso!!")
    else:
        print('\nTaxa não adicionada.')
    time.sleep(1)
    clear_console()
    print("------------------------------------------------")
    sair = input("Deseja sair do programa?(S/N) ")
    print("------------------------------------------------")
    clear_console()

# print(selectType)
# print(selectTable)
# print(nomeVendedor)
# scriptEmail()


