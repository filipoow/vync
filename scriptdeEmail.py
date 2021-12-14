import pandas as pd

df_base = pd.read_excel('base.xlsx', index_col=False)

#Procurando linhas duplicadas
linhaDuplicadas = df_base[df_base.duplicated('cnpj',keep=False)]
dadosduplicados = df_base[df_base.duplicated('cnpj')]

#Procurando linha únicas
linhaUnicas = df_base.drop_duplicates('cnpj', keep=False)
base_unica = linhaUnicas.to_dict('index')


def scriptEmailDuplicado():
    for i in dadosduplicados.index:
        #Requisitos para procura de dados...
        cnpj = dadosduplicados['cnpj'][i]

        #Configurando o layout
        nomeVendedor = dadosduplicados['nomeVendedor'][i]
        cnpj = dadosduplicados['cnpj'][i]
        origem = dadosduplicados['Origem'][i]
        seTaxa = dadosduplicados['seTaxa'][i]
        taxa = dadosduplicados['taxa'][i]
        if seTaxa.lower() == 'sim':
            #Imprimindo email
            print("------------------------------------------------")
            print("          ==Gerando Script de E-mail==          ")
            print(f"         -- Número: {len(dadosduplicados)}       ")
            print("------------------------------------------------")
            print(f"\n\n{nomeVendedor}, boa tarde!")
            print(f"\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
            #Procura de dados do tipoTabela....
            dados = df_base.loc[linhaDuplicadas['cnpj'] == cnpj]
            #Loop para 
            for i in dados.index:
                print("{} + {}".format(dados['tipoTabela'][i], dados['taxa'][i]))
            print(f"Origem {origem}")
            print("\nPor gentileza solicitar testes.")
            print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
        elif seTaxa.lower() == 'não':
             #Imprimindo email
            print("------------------------------------------------")
            print("          ==Gerando Script de E-mail==          ")
            print(f"         -- Número: {len(dadosduplicados)}       ")
            print("------------------------------------------------")
            print(f"\n\n{nomeVendedor}, boa tarde!")
            print(f"\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
            #Procura de dados do tipoTabela....
            dados = df_base.loc[linhaDuplicadas['cnpj'] == cnpj]
            #Loop para 
            for i in dados.index:
                print(dados['tipoTabela'][i])
            print(f"Origem {origem}")
            print("\nPor gentileza solicitar testes.")

def scriptEmailUnico():
    #Definindo critérios
    Identificador = 0
    for item in base_unica:
        i = item
        Identificador += 1
        #Dados para realização do scriptdeEmailUnico
        if i in base_unica:
            nomeVendedor = str(base_unica[i]['nomeVendedor'])
            cnpj = str(base_unica[i]['cnpj'])
            tipoTabela = str(base_unica[i]['tipoTabela'])
            origem = str(base_unica[i]['Origem'])
            seTaxa = str(base_unica[i]['seTaxa'])
            taxa = str(base_unica[i]['taxa'])
        else:
            print("Não foi possível encontrar os dados!")
        if seTaxa.lower() == 's':
            with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                arquivo.write("\n----------------------------------------------------------------------------------")
                arquivo.write("\n                         ==Gerando Script de E-mail==                             ")
                arquivo.write(f"\n                         -- Processando: {Identificador}/{len(base_unica)}        ")
                arquivo.write("\n----------------------------------------------------------------------------------")
                arquivo.write(f"\n{nomeVendedor}, boa tarde!")
                arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                arquivo.write(f"\n{tipoTabela} + {taxa}")
                arquivo.write(f"\nOrigem {origem}")
                arquivo.write("\n\nPor gentileza solicitar testes.")
                arquivo.write("\n\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
        elif seTaxa.lower() == 'n':
            with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                arquivo.write("\n----------------------------------------------------------------------------------")
                arquivo.write("\n                         ==Gerando Script de E-mail==                             ")
                arquivo.write(f"\n                         -- Processando: {Identificador}/{len(base_unica)}        ")
                arquivo.write("\n----------------------------------------------------------------------------------")
                arquivo.write(f"\n{nomeVendedor}, boa tarde!")
                arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                arquivo.write(f"\n{tipoTabela}")
                arquivo.write(f"\nOrigem {origem}")
                arquivo.write("\n\nPor gentileza solicitar testes.")
                arquivo.write("\n\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
        else:
            print("Não foi possível gerar os scripts, informar o desenvolvedor.")
scriptEmailDuplicado()