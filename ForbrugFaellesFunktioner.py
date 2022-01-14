import os
import pandas as pd

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
