import pandas as pd

#Importação codigos
df_codigos_vinculos = pd.read_excel('base_de_dados\codigo_ref.xlsx',dtype=object)

#Importação base de dados
base = pd.read_excel('base.xlsx')

for i in base.index:
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

    #Filtrando por tipo para que possa buscar a origem
    df_origem = df_codigos[df_codigos['IATA'] == origem]

    #Adicionando os códigos selecionados na lista
    for item in df_origem.index:
        cod_temp.append(df_origem['Cod_Tabela'][item])
    
    for element in cod_temp:
        cod.append(str(element))
    print(cod)
