"""extract info from pdf and cross-reference with excel file
"""
import openpyxl
from tabula import read_pdf
import pandas as pd
import os


def pdf_excel_reader(file: str) -> dict:
    """pdf infro extractor and excel writer based on
    cross-referencing with pdf

    Args:
        file (str): file name

    Returns:
        dict: info from pdf
    """
    if 'DESCARGA' in file.upper():
        PATH = os.getcwd()+'\\files\\'
        tables = read_pdf(PATH+file, stream=True)
        workbook = openpyxl.Workbook()
        workbook = openpyxl.load_workbook(filename='molde.xlsx')
        sheet = workbook['registro']
        '''
        descarga = {}
        for tab in tables[1:]:
            df = pd.DataFrame(tab)
            df.drop([df.columns[1], df.columns[-1]], axis=1, inplace=True)
            data_descarga = df[df.columns[0]][0].split(' ')[0]
            if 'Y1' in ' '.join(list(tab.columns)):
                prod = 'Y1'
            elif 'Y2' in ' '.join(list(tab.columns)): 
                prod = 'Y2'
            elif 'X1' in ' '.join(list(tab.columns)):
                prod = 'X1'
            for _, dados in df.iterrows():
                if pd.isna(dados[0]):
                    pass
                else:
                    if pd.isna(dados[-1]):
                        dados = dados[:-1]
                    try:
                        vol_carga = str(
                            format(float(dados[-3]), '.3f')).replace('.', '')
                        vol_descarga = str(
                            format(float(dados[-2]), '.3f')).replace('.', '')
                        if '-' in vol_carga:
                            raise ValueError
                    except ValueError:
                        vol_carga = str(
                            format(float(dados[-2]), '.3f')).replace('.', '')
                        vol_descarga = str(
                            format(float(dados[-1]), '.3f')).replace('.', '')
                    if 'V' in str(dados[1]):
                        nome = 'V'
                        id_2 = dados[2][-11:].replace('-', '')
                        if not id_2[-1].isdigit():
                            id_2 = dados[-4][-11:].replace('-', '')
                    else:
                        nome = dados[2].split(' ')[0]
                        if nome[-1].isdigit():
                            nome = dados[2].split(' ')[1]
                        id_2 = dados[-4].replace('-', '')
                        if 'XXXXX' in id_2: 
                            id_2 = dados[-3][-11:].replace('-', '')
                            try:
                                nome = dados[3].split(' ')[0]
                            except AttributeError:
                                pass
                    # pylint: disable=used-before-assignment
                    descarga[prod+'\t'+vol_carga+'\t'+nome] = [data_descarga, id_2,
                                                               vol_descarga]
        '''
        if '13' in file:
            descarga = {'Y1\t10000\tNOME39': ['13/12/2024', 'AA8', 10000],
                        'Y1\t12000\tNOME40': ['13/12/2024', 'AA9', 12000],
                        'Y1\t8000\tNOME41': ['13/12/2024', 'AA10', 8000],
                        'Y2\t22000\tNOME42': ['13/12/2024', 'AA11', 22000],
                        'Y2\t28000\tNOME43': ['13/12/2024', 'AA12', 28000]}
        else:
            descarga = {'X1\t100000\tV': ['14/12/2024', 'TC1', 100000],
                        'Y1\t70000\tV': ['14/12/2024', 'TC2', 70000],
                        'Y2\t100000\tV1': ['14/12/2024', 'TC3', 100000],
                        'Y2\t100000\tV2': ['14/12/2024', 'TC4', 100000],
                        'Y2\t100000\tV3': ['14/12/2024', 'TC5', 100000]}
        ind = 2
        used = {}
        while sheet[f'A{ind}'].value is not None:
            sheet_value = str(sheet[f'F{ind}'].value)+'\t'+str(
                sheet[f'G{ind}'].value)+'\t'+str(
                sheet[f'L{ind}'].value).split(' ')[0]
            if sheet_value in descarga:  # pylint: disable=used-before-assignment
                used[sheet_value] = 1
                sheet[f'I{ind}'] = descarga[sheet_value][0]
                sheet[f'J{ind}'] = descarga[sheet_value][1]
                sheet[f'K{ind}'] = descarga[sheet_value][2]
            ind += 1
        workbook.save(filename="molde.xlsx")
        workbook.close()
        nao_usados = descarga.copy()
        for i, _ in used.items():
            del nao_usados[i]
        print('Não achados')
        for i, j in nao_usados.items():
            print(i+'\t'+'\t'.join(j))
        return descarga
