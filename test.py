#Importação de tabela por pandas
import pandas as pd
df = pd.read_excel('codigo_ref.xlsx')
df.to_dict()

print(df)
