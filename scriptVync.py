#Importação de tabela por pandas
import pandas as pd
import pyautogui
import time
import os

#Importação Tabelas
from config import dict_package
from config import all_package4

#Importação codigos
df_codigos_vinculos = pd.read_excel('base_de_dados\codigo_ref.xlsx',dtype=object)

#Importação base de dados
base = pd.read_excel('base.xlsx')

#Declarando critérios
i = 0 

def clear_console():
    os.system('cls')

def abrindoCentral():
    #Abrindo o Cadastro Central
    pyautogui.click(x=217, y=748)

def addTaxa():
    # #Entrada de dados...
    # print("Tipo de Taxa: PACKAGE, .COM, PICKUP, RODOVIARIO, ECONOMICO, \nEXPRESSO, CORPORATE, DOC, CARGO, EMERGENCIAL")
    # tipoTaxa = input("Digite o tipo da taxa: ")
    # taxa = input("Digite a taxa: ")

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
    elif tipoTaxa.lower() == ".com":
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
    elif tipoTaxa.lower() == "rodoviario":
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
    elif tipoTaxa.lower() == "economico":
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
    elif tipoTaxa.lower() == "expresso":
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
    elif tipoTaxa.lower() == "pickup":
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
    elif tipoTaxa.lower() == "corporate":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando Corporate
        pyautogui.click(x=576, y=568)
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
    elif tipoTaxa.lower() == "cargo":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando Cargo
        pyautogui.click(x=575, y=606)
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
    elif tipoTaxa.lower() == "doc":
        #Abrindo itens
        pyautogui.click(x=694, y=455)
        time.sleep(1)
        #Selecionando DOC
        pyautogui.click(x=572, y=552)
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
    else:
        print("Não foi possível adicionar a taxa.")
    #Fechando aba de taxa..
    time.sleep(1)
    pyautogui.click(x=893, y=250)    

def main():
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

def fim():
    print("------------------------------------------------")
    print(f"            --Total de vínculos realizados: {len(base.index)}")

#Abrindo a central para o inicio dos cadastros
abrindoCentral()  

for i in base.index:
    #Definição de critérios
    Identificador = 0 

    #Declaração de dados
    nomeVendedor = str(base['nomeVendedor'][i])
    cnpj = str(base['cnpj'][i])
    tipoTabela = str(base['tipoTabela'][i])
    origem = str(base['Origem'][i])
    seTaxa = str(base['seTaxa'][i])
    tipoTaxa = str(base['tipoTaxa'][i])
    taxa = str(base['taxa'][i])
    
    #Se for colocar os códigos manualmente, adicionar antes de compilar..
    cod = []
    cod_temp = []

    #Procura da tabela por tipo
    df_codigos = df_codigos_vinculos[df_codigos_vinculos['Modalidade'] == tipoTabela]
    if origem == 'Multi':
        for i in df_codigos.index:
            cod_temp.append(df_codigos['Cod_Tabela'][i])
    else:
        #Filtrando por tipo para que possa buscar a origem
        df_origem = df_codigos[df_codigos['IATA'] == origem]

        #Adicionando os códigos selecionados na lista
        for item in df_origem.index:
            cod_temp.append(df_origem['Cod_Tabela'][item])

    for i in cod_temp:
        cod.append(str(i))

    
    # def escolhas():
    #     #Adicionando o código selecionado -- Package Balcão
    #     if tipoTabela.lower() == 'package balcão':
    #         if origem in dict_package['PackageBalcao']:
    #             itemTabela = dict_package['PackageBalcao'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 1
    #     if tipoTabela.lower() == 'package 1':
    #         if origem in dict_package['Package1']:
    #             itemTabela = dict_package['Package1'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- .COM 1 
    #     elif tipoTabela.lower() == '.com 1':
    #         if origem in dict_package['.Com1']:
    #             itemTabela = dict_package['.Com1'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 1,5 
    #     elif tipoTabela.lower() == 'package 1.5':
    #         if origem in dict_package['Package1.5']:
    #             itemTabela = dict_package['Package1.5'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- .COM 1,5
    #     elif tipoTabela.lower() == '.com 1.5':
    #         if origem in dict_package['.Com1.5']:
    #             itemTabela = dict_package['.Com1.5'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 2 
    #     elif tipoTabela.lower() == 'package 2':
    #         if origem in dict_package['Package2']:
    #             itemTabela = dict_package['Package2'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 2 Fronteira
    #     elif tipoTabela.lower() == 'package 2 fronteira':
    #         if origem in dict_package['Package2Front']:
    #             itemTabela = dict_package['Package2Front'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- .COM 2
    #     elif tipoTabela.lower() == '.com 2':
    #         if origem in dict_package['.Com2']:
    #             itemTabela = dict_package['.Com2'][origem]
    #             cod.append(itemTabela)
    #         elif origem == 'Multi':
    #             listaCom2 = dict_package['.Com2'].values()
    #             listaCom2 = list(listaCom2)
    #             for item in listaCom2:
    #                 cod.append(item)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- .COM 2 Fronteira
    #     elif tipoTabela.lower() == '.com 2 fronteira':
    #         if origem in dict_package['.Com2Front']:
    #             itemTabela = dict_package['.Com2Front'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 4
    #     elif tipoTabela.lower() == 'package 4':
    #         if origem in dict_package['Package4']:
    #             itemTabela = dict_package['Package4'][origem]
    #             cod.append(itemTabela)
    #         elif origem == 'Multi':
    #             listaPackage4 = dict_package['Package4'].values()
    #             listaPackage4 = list(listaPackage4)
    #             for item in all_package4:
    #                 cod.append(item)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- Package 4 Fronteira
    #     elif tipoTabela.lower() == 'package 4 fronteira':
    #         if origem in dict_package['Package4Front']:
    #             itemTabela = dict_package['Package4Front'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")
    #     #Adicionando o código selecionado -- .COM 2
    #     elif tipoTabela.lower() == 'pickup':
    #         if origem in dict_package['Pickup']:
    #             itemTabela = dict_package['Pickup'][origem]
    #             cod.append(itemTabela)
    #         else:
    #             print("Erro na procura do código, informar o criador.")

    main()
    # clear_console()
    print("------------------------------------------------")
    print("            ==Operações Realizadas==            ")
    print(f"               -- Identificador: {i}           ")
    print("               -- Tabela cadastrada com sucesso!")
    if seTaxa.lower() == 'sim':
        addTaxa()
        taxa = float(taxa)
        print("               -- Taxa adicionada com sucesso!")
    elif seTaxa.lower() == 'não':
        taxa = 0
        print("               -- Taxa não adicionada!")
    else:
        print("Erro ao efetuar a ação, entrar em contato com o criador.")
fim()