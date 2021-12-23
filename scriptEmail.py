import pandas as pd
import time

#Importação da base de dados
df_base = pd.read_excel('base.xlsx',dtype=object)

#Criando uma listagem dos vendedores
lista_vendedores = list(set(df_base['nomeVendedor'].values.tolist()))

#Novo df
base_temporaria = pd.DataFrame()

def scriptEmail():
    for vendedor in lista_vendedores:
        base_temporaria = df_base[df_base['nomeVendedor'] == vendedor]
        lista_cnpj = list(set(base_temporaria['cnpj'].values.tolist()))
        for cnpj in lista_cnpj:
            base_por_cnpj = base_temporaria[base_temporaria['cnpj'] == cnpj]
            nomeVendedor = vendedor
            cnpj = cnpj
            for i in base_por_cnpj.index:
                origem = base_por_cnpj['Origem'][i]
                seTaxa = base_por_cnpj['seTaxa'][i]
                taxa = base_por_cnpj['taxa'][i]
            if seTaxa.lower() == 'sim' and len(base_por_cnpj['tipoTabela']) > 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabelas:")
                    for i in base_por_cnpj.index:
                        arquivo.write("\n{} + {}".format(base_por_cnpj['tipoTabela'][i],base_por_cnpj['taxa'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
                    arquivo.write("\n\nAs taxas foram cadastradas no sistema devendo ser inserida manualmente na emissão.")
            elif seTaxa.lower() == 'sim' and len(base_por_cnpj['tipoTabela']) == 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                    for i in base_por_cnpj.index:
                        arquivo.write("\n{} + {}".format(base_por_cnpj['tipoTabela'][i],base_por_cnpj['taxa'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
                    arquivo.write("\n\nA taxa foi cadastrada no sistema devendo ser inserida manualmente na emissão.")
            elif seTaxa.lower() == 'não' and len(base_por_cnpj['tipoTabela']) > 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foram cadastradas no CNPJ {cnpj} as tabela:")
                    for i in base_por_cnpj.index:
                        arquivo.write("\n{}".format(base_por_cnpj['tipoTabela'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
            elif seTaxa.lower() == 'não' and len(base_por_cnpj['tipoTabela']) == 1:
                with open(f'scriptEmail\{cnpj}.txt', 'w') as arquivo:
                    #Imprimindo email
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write("\n          ==Gerando Script de E-mail==          ")
                    arquivo.write(f"\n         -- Número: {len(lista_cnpj)}       ")
                    arquivo.write("\n------------------------------------------------")
                    arquivo.write(f"\n\n{nomeVendedor}, boa tarde!")
                    arquivo.write(f"\n\nConforme solicitado foi cadastrada no CNPJ {cnpj} a tabela:")
                    for i in base_por_cnpj.index:
                        arquivo.write("\n{}".format(base_por_cnpj['tipoTabela'][i]))
                    arquivo.write(f"\nOrigem {origem}")
                    arquivo.write("\n\nPor gentileza solicitar testes.")
            else:
                print("Não foi possível gerar o email.")
def gerandoEmail():
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
gerandoEmail()
