# from openpyxl.workbook import defined_name
from openpyxl import load_workbook

wb = load_workbook(filename=
                   'C:\\SRA\\LIFE-ForFit Financial Reporting\\1. SaltenLangsø\\1. Financial Report\\'
                   'TestSaltenLangso.xlsx')

definedcelle = {'navn': None, 'fane': None, 'celle': None}
liste_af_definedceller = []

for dn in wb.defined_names.definedName:
    definedcelle['navn'] = dn.name
    definedcelle['fane'] = dn.attr_text.split('!')[0]
    try:
        definedcelle['celle'] = dn.attr_text.split('!')[1]
    except IndexError:
        definedcelle['celle'] = 'Ingen celle henvisning'  # dn.attr_text.split('!')[0]

    liste_af_definedceller.append(definedcelle.copy())

for henvis in liste_af_definedceller:
    for key, value in henvis.items():
        print(f'{key}: {value}')

    print()
