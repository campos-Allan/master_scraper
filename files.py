"""file selector and formatter
"""

import os
import pandas as pd
from tabula import read_pdf


def sem_data(file: str) -> str:
    """putting date in the file name

    Args:
        file (str): file name

    Returns:
        str: exported date from file
    """
    if 'SALDO' in file.upper():
        df = pd.DataFrame(read_pdf(file, stream=True)[0])
        return df[df.columns[1]][0].replace('/', '.')
    elif 'DESCARGA' in file.upper():
        df = pd.DataFrame(read_pdf(file, stream=True)[1])
        return df[df.columns[0]][0].split(' ')[0].replace('/', '.')


def trash(files):
    """take files off the folder
    """
    for file in files:
        if 'molde' in file:
            pass
        else:
            os.rename(f'{os.getcwd()}\\files\\{file}',
                      f'{os.getcwd()}\\trash\\{file}')


PATH = os.getcwd()+'\\files\\'
FILES_PDF = [f for f in os.listdir(PATH) if 'pdf' in f or 'PDF' in f]
FILES_EXCEL = [f for f in os.listdir(PATH) if 'xlsx' in f]
'''
for i in FILES_PDF:
    if any(char.isdigit() for char in i):
        pass
    else:
        if '(SN)' in i:
            pass
        else:
            dia = sem_data(i)
            os.rename(i, f'{i[:-4]} {dia}.pdf')
FILES_PDF = [f for f in os.listdir(PATH) if 'pdf' in f or 'PDF' in f]
'''
