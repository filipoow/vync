import pandas as pd

df_base = pd.read_excel('base.xlsx', index_col=False)

print(df_base)

for index, row in df_base.iterrows():
    print(row['cnpj'],row['tipoTabela'])
    