import pandas as pd
import time

df_base = pd.read_excel('base.xlsx',dtype=object)

lista_vendedores = list(set(df_base['nomeVendedor'].values.tolist()))

df_temp = pd.DataFrame()

def scriptEmail():
    for vendedor in lista_vendedores:
        df_temp = df_base[df_base['nomeVendedor'] == vendedor]
        lista_cnpj = list(set(df_temp['cnpj'].values.tolist()))
        for cnpj in lista_cnpj:
            df_cnpj = df_temp[df_temp['cnpj'] == cnpj]
            nomeVendedor = vendedor
            cnpj = cnpj
            for i in df_cnpj.index:
                origem = df_cnpj['Origem'][i]
                seTaxa = df_cnpj['seTaxa'][i]
                taxa = df_cnpj['taxa'][i]
            if seTaxa.lower() == 'sim' and len(df_cnpj['tipoTabela']) > 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabelas:")
                    for i in df_cnpj.index:
                        arquivo.write("\n{} + {}".format(df_cnpj['tipoTabela'][i],df_cnpj['taxa'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
                    arquivo.write("\n\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
            elif seTaxa.lower() == 'sim' and len(df_cnpj['tipoTabela']) == 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                    for i in df_cnpj.index:
                        arquivo.write("\n{} + {}".format(df_cnpj['tipoTabela'][i],df_cnpj['taxa'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
                    arquivo.write("\n\nA taxa foi cadastrada no sistema devendo ser inserida manualmente na emissão.")
            elif seTaxa.lower() == 'não' and len(df_cnpj['tipoTabela']) > 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
                    for i in df_cnpj.index:
                        arquivo.write("\n{}".format(df_cnpj['tipoTabela'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
            elif seTaxa.lower() == 'não' and len(df_cnpj['tipoTabela']) == 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                    for i in df_cnpj.index:
                        arquivo.write("\n{}".format(df_cnpj['tipoTabela'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
            else:
                print("Não foi possível gerar o email.")
print("\n-----------------------------")
print("Gerando Script de E-mail....")
scriptEmail()
time.sleep(1)
print("\nEmails Únicos - OK")
time.sleep(1)
print("Emails Duplos - OK")
time.sleep(1)
print("\n-----------------------------")
print("Aplicações Finalizadas!")


