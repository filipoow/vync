import pandas as pd

df_base = pd.read_excel('base.xlsx', index_col=False)

linhaDuplicadas = df_base[df_base.duplicated('cnpj',keep=False)]

print(linhaDuplicadas)
dadosduplicados = df_base[df_base.duplicated('cnpj')]
print(dadosduplicados)
for i in dadosduplicados.index:
    print(i)
    cnpj = dadosduplicados['cnpj'][i]

    nomeVendedor = dadosduplicados['nomeVendedor'][i]
    origem = dadosduplicados['Origem'][i]
    seTaxa = dadosduplicados['seTaxa'][i]
    taxa = dadosduplicados['taxa'][i]
    if seTaxa.lower() == 'sim':
        print(cnpj,nomeVendedor,origem,seTaxa,taxa)
        dados = df_base.loc[linhaDuplicadas['cnpj'] == cnpj]
        for i in dados.index:
            print("{} + {}".format(dados['tipoTabela'][i], dados['taxa'][i]))

    
# for i in dadosduplicados.index:
#     #Requisitos para procura de dados...
#     cnpj = dadosduplicados['cnpj'][i]

#     #Configurando o layout
#     nomeVendedor = dadosduplicados['nomeVendedor'][i]
#     cnpj = dadosduplicados['cnpj'][i]
#     origem = dadosduplicados['Origem'][i]
#     seTaxa = dadosduplicados['seTaxa'][i]
#     taxa = dadosduplicados['taxa'][i]
#     if seTaxa.lower() == 'sim':
#         #Imprimindo email
#         print("------------------------------------------------")
#         print("          ==Gerando Script de E-mail==          ")
#         print(f"         -- Número: {len(dadosduplicados)}       ")
#         print("------------------------------------------------")
#         print(f"\n\n{nomeVendedor}, boa tarde!")
#         print(f"\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
#         #Procura de dados do tipoTabela....
#         dados = df_base.loc[linhaDuplicadas['cnpj'] == cnpj]
#         #Loop para 
#         for i in dados.index:
#             print("{} + {}".format(dados['tipoTabela'][i], dados['taxa'][i]))
#         print(f"Origem {origem}")
#         print("\nPor gentileza solicitar testes.")
#         print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
#     elif seTaxa.lower() == 'não':
#          #Imprimindo email
#         print("------------------------------------------------")
#         print("          ==Gerando Script de E-mail==          ")
#         print(f"         -- Número: {len(dadosduplicados)}       ")
#         print("------------------------------------------------")
#         print(f"\n\n{nomeVendedor}, boa tarde!")
#         print(f"\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
#         #Procura de dados do tipoTabela....
#         dados = df_base.loc[linhaDuplicadas['cnpj'] == cnpj]
#         #Loop para 
#         for i in dados.index:
#             print(dados['tipoTabela'][i])
#         print(f"Origem {origem}")
#         print("\nPor gentileza solicitar testes.")