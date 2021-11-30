#Importação de tabela por pandas
import pandas as pd
df = pd.read_excel('codigo_ref.xlsx')
print(df)

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



