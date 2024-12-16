"""extracts info from pdf
"""

from datetime import datetime
import pandas as pd
from tabula import read_pdf
from pypdf import PdfReader
import os


def pdf_reader(file: str):
    """pdf info extractor for each type of file

    Args:
        file (str): file name

    Returns:
        _type_: dataframe or dictionary with info
    """
    # pylint: disable=unsubscriptable-object
    if 'ESTOQUE' in file.upper():
        PATH = os.getcwd()+'\\files\\'
        estoque = pd.DataFrame(read_pdf(PATH+file, stream=True)[0])
        estoque.columns = estoque.iloc[0]
        estoque = estoque.drop(0)
        # The original pdf had a defect of joining some columns when read, so this next line was necessary in the real file
        # estoque.loc[:, 'Entradas'] = estoque[estoque.columns[0]].str.split(' ', expand=True)[2]
        estoque = estoque[['Entradas', 'Est. Final']].loc[[2, 8, 12], :]
        estoque = estoque.rename(index={2: 'X1', 8: 'Y2', 12: 'Y1'}, columns={
            'Entradas': 'entrada_', 'Est. Final': 'saldo_'})
        estoque['entrada_'] = estoque['entrada_'].map(
            lambda x: x.replace('.', '')).astype(int)
        return estoque
    elif 'MOV' in file.upper():
        PATH = os.getcwd()+'\\files\\'
        mov = {}
        ano = datetime.now().strftime('%Y')
        reader = PdfReader(PATH+file)
        number_of_pages = len(reader.pages)
        for j in range(0, number_of_pages):
            page = reader.pages[j]
            text = page.extract_text()
            '''THIS PART WOULD READ THE PDF FILE AND SELECT THE PART WITH DATA TO SCRAPE
            if 'Veículo' in text:
                info = pd.DataFrame(text.split('\n'))
                corte1 = info.loc[info[0] == 'Saídas'].index[0].tolist()
                while info.iloc[corte1][0].find(ano) == -1:
                    corte1 += 1
                corte2 = info.loc[info[0] ==
                                  'CONTROLE E EXPEDIÇÃO DE COMBUSTÍVEIS'].index[0].tolist()
                '''
            # trecho = pd.Series(info.iloc[corte1:corte2][0])
            if j in (0, 3):
                trecho = pd.Series(
                    ['13/12/2024', '/144444', 'SSSSSSS1', 'NOME RANDOM 2', '60.100', '60.000'])
                trecho.name = 'X1'
            elif j in (1, 4):
                trecho = pd.Series(
                    ['13/12/2024', '/133332', 'V', 'V', '100.100', '100.000',
                        '13/12/2024', '/133333', 'V', 'V', '100.100', '100.000',
                        '13/12/2024', '/133334', 'V', 'V', '100.100', '100.000'])
                trecho.name = 'Y1'
            elif j in (2, 5):
                trecho = pd.Series(
                    ['13/12/2024', '/133335', 'V', 'V', '100.100', '100.000',
                        '13/12/2024', '/133336', 'V', 'V', '100.100', '100.000'])
                trecho.name = 'Y2'
            for l in range(0, len(trecho)+1, 6):
                if l == 0:
                    pass
                else:
                    reg = trecho.iloc[l-6:l].to_list()
                    reg.append(trecho.name)
                    num = reg.pop(1)
                    mov[num] = reg
        return mov
    elif 'SALDO' in file.upper():
        PATH = os.getcwd()+'\\files\\'
        saldo = pd.DataFrame(read_pdf(PATH+file, stream=True)[0])
        if saldo[saldo.columns[7]].isna().any():
            saldo = saldo[[saldo.columns[8], saldo.columns[-1]]]
        else:
            saldo = saldo[[saldo.columns[7], saldo.columns[-1]]]
        saldo.drop(0, inplace=True)
        saldo[saldo.columns[0]] = saldo[saldo.columns[0]].apply(
            lambda x: float(str(x).replace('0 0', '0')))
        saldo[saldo.columns[0]] = saldo[saldo.columns[0]].apply(lambda x: format(
            x, '.3f')).astype(str).replace('0.000', '0')
        change_b, change_c = saldo.iloc[1], saldo.iloc[2]
        saldo.iloc[1] = change_c
        saldo.iloc[2] = change_b
        saldo[saldo.columns[0]] = saldo[saldo.columns[0]].map(
            lambda x: x.replace('.', '')).astype(int)
        saldo = saldo.rename(index={3: 'Y1', 1: 'X1', 2: 'Y2'}, columns={
            saldo.columns[0]: 'mercado_', saldo.columns[1]: 'saldo_'})
        return saldo
