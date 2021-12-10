import pandas as pd

df_base = pd.read_excel('base.xlsx', index_col=False)

#Printando DataFrame da Base de Dados
print(df_base)

#Procurando linhas duplicadas
linhaDuplicadas = df_base[df_base.duplicated('cnpj',keep=False)]
print("-----------------------------------------------------------------------")
print("                                           --Buscando linhas duplicadas")
print(linhaDuplicadas)

linhaUnicas = df_base.drop_duplicates('cnpj', keep=False)

#Buscando linhas únicas
print("-----------------------------------------------------------------------")
print("                                               --Buscando linhas únicas")
print(linhaUnicas)

# Convertendo df para dicionário
base_duplicada = linhaDuplicadas.to_dict('index')
base_unica = linhaUnicas.to_dict('index')

#Printando Dicionário
print("-----------------------------------------------------------------------")
print("                                      --Dicionário de linhas duplicadas")
print(base_duplicada)
print("-----------------------------------------------------------------------")
print("                                          --Dicionário de linhas únicas")
print(base_unica)


def scriptEmailDuplicado():
    for item in base_duplicada:
        i = item 
        if i in base_duplicada:
            nomeVendedor = str(base_duplicada[i]['nomeVendedor'])
            cnpj = str(base_duplicada[i]['cnpj'])
            for dupli in base_duplicada[i]['cnpj']:
                dupli = str(dupli)
                listaTabelas = []
                listaTabelas.append(str(base_duplicada['tipoTabela']))
            origem = str(base_duplicada[i]['Origem'])
            seTaxa = str(base_duplicada[i]['seTaxa'])
            taxa = str(base_duplicada[i]['taxa'])
        else:
            print("Não foi possível encontrar os dados!")  
        if seTaxa.lower() == 's':
            print("------------------------------------------------")
            print("          ==Gerando Script de E-mail==          ")
            print(f"         -- Número: {len(base_duplicada)}       ")
            print("------------------------------------------------")
            print(f"\n\n{nomeVendedor}, boa tarde!")
            print(f"\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
            for item in listaTabelas:
                print(f"{item} + {taxa}")
            print(f"Origem {origem}")
            print("\nPor gentileza solicitar testes.")
            print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
        elif seTaxa.lower() == 'n':
            print("------------------------------------------------")
            print("          ==Gerando Script de E-mail==          ")
            print(f"         -- Número: {len(base_duplicada)}       ")
            print("------------------------------------------------")
            print(f"\n\n{nomeVendedor}, boa tarde!")
            print(f"\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
            for item in listaTabelas:
                print(item)
            print(f"Origem {origem}")
            print("\nPor gentileza solicitar testes.")
            print("\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")

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
                arquivo.write(f"\n                        -- Processando: {Identificador}/{len(base_unica)}        ")
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
                arquivo.write(f"\n                        -- Processando: {Identificador}/{len(base_unica)}        ")
                arquivo.write("\n----------------------------------------------------------------------------------")
                arquivo.write(f"\n{nomeVendedor}, boa tarde!")
                arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                arquivo.write(f"\n{tipoTabela}")
                arquivo.write(f"\nOrigem {origem}")
                arquivo.write("\n\nPor gentileza solicitar testes.")
                arquivo.write("\n\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
        else:
            print("Não foi possível gerar os scripts, informar o desenvolvedor.")
scriptEmailUnico()
