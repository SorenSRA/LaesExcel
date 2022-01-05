import pandas as pd

kolonner = [0, 1, 2, 3, 4]
kol_names = ['A', 'B', 'C', 'D', 'E']

df = pd.read_excel('TestSaltenLangso.xlsx',
                   sheet_name='Individual Cost Statement',
                   usecols=kolonner,
                   names=kol_names,
                   header=None)

celle_dic = {
    'valuta': (37, 'E'),
    'partner': (5, 'B'),
    'periode_to': (3, 'E'),

}
key = 'valuta'
celle = celle_dic[key]
print(f"Valuta: {df.loc[celle_dic['valuta']]} \n"
      f"Partner: {df.loc[celle_dic['partner']]} \n"
      f"Periode To: {df.loc[celle_dic['periode_to']]}")
