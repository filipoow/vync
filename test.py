#Importação de tabela por pandas
# import pandas as pd
# df = pd.read_excel('base.xlsx', index_col=False)
# print(df)

# # Convert the DataFrame to Series

# base = df.to_dict('index')
# base_values = base.items()

# new_base = {str(key): str(value) for key, value in base_values}
# print("Df to dict: ")

# print(base)
# print("--------------------------------- \nNova base como string: ")
# print(new_base);

cod = []

dict_package = {
    'PackageBalcao': {
        'SAO': ''
    },
    'Package1': {
        'SAO': '9727', 'SSA': '12618', 'SLZ': '12932', 'FOR': '11032', 'RIO': '9409','GYN': '11298', 'REC': '9980', 'BHZ': '9716',
        'CWB': '9694', 'BSB': '10167', 'FLN': '9327', 'POA': '13088', 'MCZ': '11947', 'AJU': '10859', 'VIX': '9729'
    },
    '.Com1': {
        'BHZ': '9717', 'SSA': '12619', 'SAO': '9728', 'SLZ': '12933', 'RIO': '9410', 'FOR': '9479', 'GYN': '10060', 'REC': '9741',
        'CWB': '9695', 'BSB': '10170', 'FLN': '9708', 'POA': '13090', 'MCZ': '11948', 'AJU': '10860','VIX':'9730'
    },
    'Package1.5': {
        'SAO': '10640'
    },
    '.Com1.5': {
        'SAO': '10883', 'BHZ': '12977'
    },
    'Package2': {
        'SAO': '10461', 'GYN': '9856','SSA': '11151','BHZ': '11108','RIO': '10126'
    },
    '.Com2': {
        'RIO': '12019', 'SAO': '10462', 'MAO': '12482', 'CGB': '12591', 'MCZ': '12801', 'NAT': '12866', 'BHZ': '12021', 'JPA': '12020',
        'GYN':'12025', 'VIX':'10216', 'AMK': '13034', 'SSA': '11152', 'ARU': '13035','ATB':'13036','BAU':'13037','BNU':'13039','BTV':'13038','CCM':'13040',
        'CGB':'13041','CPQ':'13042','CPV':'13043','EXT':'13044','FLN':'12023','FOR':'12024'
    },
    'Package4':{
        'SAO': '12436', 'SSA': '12772','BHZ': '12724','VIX':'12780'
    },
    'Package2Front': {
        'RIO': '12874', 'SAO': '11255'
    },
    'Package4Front': {
        'SAO': '12612'
    },
    'Pickup': {
        'SAO': '11532'
    }
}
typeTable = input("Digite o tipo da tabela: ")
#Selecionando Package Balcão
if typeTable == 'Package Balcão':
    opt_table = input("Busque a tabela: ")
    if opt_table in dict_package['Package1']:
        print("Encontrado!")
        tableItem = dict_package['Package1'][opt_table]
        cod.append(tableItem)
#Selecionando Package 1
if typeTable == 'Package 1':
    opt_table = input("Busque a tabela: ")
    if opt_table in dict_package['Package1']:
        print("Encontrado!")
        tableItem = dict_package['Package1'][opt_table]
        cod.append(tableItem)
#Selecionando Package 2
if typeTable == 'Package 2':
    opt_table = input("Busque a tabela: ")
    if opt_table in dict_package['Package2']:
        print("Encontrado!")
        tableItem = dict_package['Package2'][opt_table]
        cod.append(tableItem)
if typeTable == 'Package 4':
    opt_table = input("Digite a origem da tabela: ")
    if opt_table in dict_package['Package4']:
        tableItem = dict_package['Package4'][opt_table]
        cod.append(tableItem)
    elif opt_table == 'Multi':
        listaPackage4 = dict_package['Package4'].values()
        listaPackage4 = list(listaPackage4)
        for item in listaPackage4:
            cod.append(item)

print(cod)