import os
import pandas as pd
import numpy
from openpyxl.utils.dataframe import dataframe_to_rows

pathbase = 'C:\\SRA\\LIFE-ForFit Financial Reporting\\'
pathspecff = '\\1. Financial Report'

euro_til_dkr = 7.4393
dkr = 'Dkr'


def get_partner_input_wb(partn):
    path = pathbase + partn + pathspecff
    print('2 ' + str(os.listdir(path)))
    wb_found = False
    for file in os.listdir(path):
        if (not wb_found) and (os.path.splitext(file)[1] == '.xlsx'):
            excel_fil_navn = str(os.path.join(path, file))
            print('3' + '  ' + excel_fil_navn)
            wb_found = True

    if wb_found:
        return True, str(os.path.join(path, file))
    else:
        return False, "Excel-fil ikke fundet"


def get_valuta_kor(input_val, output_valuta):
    if output_valuta == dkr:  # output_valuta er Dkr
        if input_val == dkr:  # input_valuta er Dkr
            val_kor = 1
        else:  # input_valuta er EURO
            val_kor = euro_til_dkr
    else:  # output_valuta er EURO
        if input_val == dkr:
            val_kor = 1.0 / euro_til_dkr
        else:
            val_kor = 1.0
    return val_kor


def get_valuta(wb):
    valuta_sh = 'Individual Cost Statement'
    euro_kol = 'E'
    euro_rk = 37
    valuta_df = pd.read_excel(wb,
                              sheet_name=valuta_sh,
                              usecols=euro_kol,
                              header=None,
                              names=[euro_kol])

    return valuta_df.loc[euro_rk, euro_kol]


# udtraek skal være enten 'Cost', 'Income', 'Actions'
def get_input_df(wb, udtraek='Cost'):
    if udtraek == 'Cost':
        cols_laes = 'B'
        rk_start_skip = 10
        rk_laes_antal = 12
        laes_sh = 'Individual Cost Statement'
        cols_navn = ['Indlaes']
    elif udtraek == 'Income':
        cols_laes = 'E'
        rk_start_skip = 10
        rk_laes_antal = 4
        laes_sh = 'Individual Cost Statement'
        cols_navn = ['Indlaes']
    else:   # udtraek == 'Actions':
        cols_laes = 'C:Z'
        rk_start_skip = 3
        rk_laes_antal = 10
        laes_sh = 'BudgetDisponering'
        cols_navn = ['C', 'D', 'E']

    return pd.read_excel(wb,
                         sheet_name=laes_sh,
                         skiprows=rk_start_skip,
                         usecols=cols_laes,
                         nrows=rk_laes_antal,
                         header=None,
                         names=cols_navn,
                         keep_default_na=False,
                         )


def get_input_eurodk(df, input_val):
    df.insert(1, 'Euro', 0)
    df.insert(2, 'Dkr', 0)
    for ix in df.index:
        if type(df.at[ix, 'Indlaes']) in (int, float, complex, numpy.float64):
            df.at[ix, 'Euro'] = df.at[ix, 'Indlaes'] * get_valuta_kor(input_val, 'Euro')
            df.at[ix, 'Dkr'] = df.at[ix, 'Euro'] * euro_til_dkr

    return df.drop(columns=['Indlaes'])


def get_output_fane(wb, sh_name):
    output_skabelon_sh = 'StdSheet'
    if sh_name not in wb.sheetnames:
        wb.copy_worksheet(wb[output_skabelon_sh]).title = sh_name

    return wb[sh_name]


def kopier_df_excel(outputfane, input_dataframe, udtraek='Cost'):
    if udtraek == 'Cost':
        kol_korrektion = 2

    elif udtraek == 'Income':
        kol_korrektion = 5

    else:   # udtraek == 'Actions':
        kol_korrektion = 5

    rk = 11
    for row_str in dataframe_to_rows(input_dataframe, index=False, header=False):
        for kol, cel in enumerate(row_str):
            outputfane.cell(row=rk, column=kol + kol_korrektion).value = int(cel)
        rk += 1

    return
