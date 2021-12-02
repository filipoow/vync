#Importação de tabela por pandas
# import pandas as pd
# df = pd.read_excel('codigo_ref.xlsx')
# print(df)

#TESTE PARA CADASTRO DE CÓDIGOS..............................................

# Package1 = {'SAO': 9727, 'SSA': 12618, 'SLZ': 12932, 'FOR': 11032, 'RIO': 9409,'GYN': 11298, 'REC': 9980, 'BHZ': 9716,
# 'CWB': 9694, 'BSB': 10167, 'FLN': 9327, 'POA': 13088, 'MCZ': 11947, 'AJU': 10859, 'VIX': 13530}
# cod = []
# numIdentificador = 0 
# qntTabelas = int(input("Quantas tabelas deseja cadastrar? "))
# qntTabelas = qntTabelas + 1
# while numIdentificador < qntTabelas:

# selectType = input("Digite o tipo da tabela: ")
# if selectType == 'Package 1':
#     selectTable = input("Digite a tabela a Origem: ")
#     if selectTable in Package1:
#         PackageItem = Package1[selectTable]
#         cod.append(PackageItem)
# print(cod)

# def scriptEmail():
#     #Entrada de dados..
#     print("\n-------------------------------------")
#     print("            Script Email             ")
#     print("-------------------------------------")
#     # print("{}"", boa tarde!".format(nomeVendedor))
#     print("\nConforme solicitado foram cadastradas no CNPJ {} as tabelas:".format(cnpj))
#     print("{} + {}".format(selectType,taxa))
#     print("Origem {}".format(selectTable))
#     print("\nPor gentileza solicitar testes.")
#     print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")

# cnpj = input("Digite o cnpj: ")
# selectType = input("Digite o tipo da tabela: ")
# selectTable = input("Digite a Origem: ")
# taxa = input("Digite a taxa: ")
# scriptEmail()

# a_list = []

# qtn = int(input("Quantidade de vínculos: "))
# while len(a_list) < qtn:
#     item = input("Digite o código: ")
#     a_list.append(item)

# print(a_list)

itemsEmail = {}

def requisitosEmail():
    tipoTabela = input("Digite o tipo da tabela: ")
    chave = "Tabela"
    itemsEmail.setdefault(chave, [])
    itemsEmail[chave].append(tipoTabela)

    
    listadeItens = itemsEmail.values()

    listadeItens = list(listadeItens)


    print(listadeItens)
    
    #Pegando itens separados..
    for i in listadeItens:
        print(i)
        i.split([''])
 
requisitosEmail()
