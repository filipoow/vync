import pandas as pd
import pyautogui
import time
import os

#Importação Tabelas
from config import dict_package
from config import all_package4
from scriptVync import abrindoCentral

df = pd.read_excel('base.xlsx', index_col=False)
# Convertendo df para dicionário
base = df.to_dict('index')

abrindoCentral()
for i in range(len(base)):
    print(i)
    nomeVendedor = str(base[i]['nomeVendedor'])
    cnpj = str(base[i]['cnpj'])
    tipoTabela = str(base[i]['tipoTabela'])
    origem = str(base[i]['Origem'])
    seTaxa = str(base[i]['seTaxa'])
    tipoTaxa = str(base[i]['tipoTaxa'])
    taxa = str(base[i]['taxa'])

    cod = []
    
    if tipoTabela.lower() == '.com 2':
        if origem in dict_package['.Com2']:
            itemTabela = dict_package['.Com2'][origem]
            cod.append(itemTabela)

    print(cod)