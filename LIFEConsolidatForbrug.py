import pandas as pd
import ForbrugFaellesFunktioner as fff
from openpyxl import load_workbook

from ForFitPartneroversigtLocal import partnerliste


output_wb_skab = 'C:\\SRA\\LIFE-ForFit Financial Reporting\\' \
                    '999. OkonomiOpfolgning\\5. Forbrug\\LifeForFitForbrugConsolSkabelon.xlsx'
output_wb = 'C:\\SRA\\LIFE-ForFit Financial Reporting\\' \
                '999. OkonomiOpfolgning\\5. Forbrug\\LifeForFitForbrugColsol2021Q3.xlsx'


output_valuta = 'Dkr'  # Enten 'Euro' eller 'Dkr'
euro_til_dkr = 7.4365

dkr = 'Dkr'

# åbn Excell-fil med standardtekst/-tabel
wb_output = load_workbook(output_wb_skab)

# opret dataframe til løbende summation af de enkelte partneres indlæste dataframes
ialt_df_cost = pd.DataFrame()
ialt_df_income = pd.DataFrame()

for key, partner in partnerliste.items():
    input_wb_exsist, input_wb = fff.get_partner_input_wb(partner)
    if not input_wb_exsist:
        print(f'{key} : Finansiel rapport ikke fundet')
        continue
    else:
        print(f'{key} : Finansiel rapport fundet')

    input_df_cost = fff.get_input_df(input_wb, 'Cost')
    print(input_df_cost)
    input_df_income = fff.get_input_df(input_wb, 'Income')
    print(input_df_income)
    input_df_cost = fff.get_input_eurodk(input_df_cost, fff.get_valuta(input_wb))
    print(input_df_cost)
    input_df_income = fff.get_input_eurodk(input_df_income, fff.get_valuta(input_wb))
    print(input_df_income)

    sh_outputfane = fff.get_output_fane(wb_output, key)

    fff.kopier_df_excel(sh_outputfane, input_df_cost, 'Cost')
    fff.kopier_df_excel(sh_outputfane, input_df_income, 'Income')

    # løbende summering af input_dataframes for hver enkelt partner
    ialt_df_cost = ialt_df_cost.add(input_df_cost, fill_value=0)
    ialt_df_income = ialt_df_income.add(input_df_income, fill_value=0)

sh_outputfane = fff.get_output_fane(wb_output, 'I alt')
fff.kopier_df_excel(sh_outputfane, ialt_df_cost, 'Cost')
fff.kopier_df_excel(sh_outputfane, ialt_df_income, 'Income')
#
# # gem Excell-fil med de kopierede forbrugs-/budgettal
wb_output.save(output_wb)
