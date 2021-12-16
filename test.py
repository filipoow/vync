import pandas as pd
from pandas.core.indexes import base

#Importação codigos
df_codigos_vinculos = pd.read_excel('base_de_dados\codigo_ref.xlsx',dtype=object)

#Importação base de dados
base = pd.read_excel('base.xlsx')
print(base)
# Convertendo df para dicionário

#Declarando critérios
i = 0 

for i in base.index:
    nomeVendedor = str(base['nomeVendedor'][i])
    cnpj = str(base['cnpj'][i])
    tipoTabela = str(base['tipoTabela'][i])
    origem = str(base['Origem'][i])
    seTaxa = str(base['seTaxa'][i])
    tipoTaxa = str(base['tipoTaxa'][i])
    taxa = str(base['taxa'][i])
    
    #Se for colocar os códigos manualmente, adicionar antes de compilar..
    cod = []

    df_codigos = df_codigos_vinculos[df_codigos_vinculos['Modalidade'] == tipoTabela]
    print(df_codigos)
    df_origem = df_codigos[df_codigos['IATA'] == origem]
    print(df_origem)
    for i in df_origem.index:
        print(df_origem['Cod_Tabela'][i])