"""pdf and excel scraping
"""
from datetime import datetime
from datetime import timedelta
import warnings
import shutil
import pandas as pd
import openpyxl
from var import registro, T1, T2, R2, S1, R1
from files import FILES_PDF, FILES_EXCEL, PATH, trash
from T1_reader import pdf_reader
from T1_insert_reader import pdf_excel_reader
from R1_S1_insert_reader import excel_reader
from S1_insert_reader import S1_excel_reader
warnings.filterwarnings("ignore")
# pylint: disable=used-before-assignment


def transformation(data, df, indexer):
    """formatting dataframe

    Args:
        data (dataframe): df with info
        df (dataframe): data formatted
        indexer (str): date in the file

    Returns:
       dataframe: df with correct formatting
    """
    for item in range(0, len(data)):
        row = data.iloc[item]
        row.index = row.index+row.name
        row = pd.DataFrame(row).transpose()
        row = row.rename(index={row.index[0]: indexer})
        df = pd.concat([df, row], axis=0, ignore_index=False)
    return df


def insertion(df_main, df_values, column, type_1, product):
    """inserting values per type of transportation, product and date

    Args:
        df_main (dataframe): dataframe to insert info
        df_values (dataframe): dataframe with the info
        column (str): column name
        type_1 (str): type of transportation
        product (str): type of product
    """
    try:
        df_main.loc[df_main.index == df_values['data_carga'][0], column] = df_values.loc[
            df_values['tipo'] == type_1][df_values['prod'] == product]['vol_carga'].reset_index(drop=True)[0]
    except KeyError:
        pass


def excel_writer(sheetname, df):
    """putting info into a control spreadsheet

    Args:
        sheetname (str)
        df (dataframe)
    """
    workbook = openpyxl.Workbook()
    workbook = openpyxl.load_workbook(filename='molde.xlsx')
    ind = 1
    sheet = workbook[sheetname]
    while sheet[f'A{ind}'].value is not None:
        ind += 1
    workbook.close()
    writer = pd.ExcelWriter(
        path='molde.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    if sheetname in ('T1', 'T2', 'R2', 'R1', 'S1'):
        df.to_excel(writer, sheet_name=sheetname,
                    header=False, index=True, startrow=ind-1)
    else:
        df.to_excel(writer, sheet_name=sheetname,
                    header=False, index=False, startrow=ind-1)
    writer.close()

try:
    shutil.rmtree(PATH+'\\trash')
except Exception as e:
    print(e)
for i in FILES_PDF:
    result = pdf_reader(i)
    if 'ESTOQUE' in i.upper():
        dia = (i.split(' ')[1][:-3]+(datetime.now() -
               timedelta(days=1)).strftime('%Y')).replace('.', '/')
        T1 = transformation(result, T1, dia)
        T1 = T1.groupby(T1.index).sum()
    elif 'MOV' in i.upper():
        for o_id, k in result.items():
            MODAL = 'MODAL A'
            if k[1] == 'V':
                k[1] = o_id
                MODAL = 'MODAL B'
            k[-2] = int(k[-2].replace('.', ''))
            registro.loc[o_id] = [k[0], 'T1', 'T2', MODAL,
                                  k[1], k[-1], k[-2], '', '', '', '', k[2]]
        regis = registro.copy()
        carga = regis.groupby(['data_carga', 'tipo', 'prod'])[
            'vol_carga'].sum().reset_index()
        ultimo = carga['data_carga'][len(carga)-1]
        carga = carga[carga['data_carga'] == ultimo].reset_index(drop=True)
        columns = ['saida_MODAL A_X1', 'saida_MODAL A_Y2', 'saida_MODAL A_Y1', 'saida_MODAL B_X1',
                   'saida_MODAL B_Y2', 'saida_MODAL B_Y1']
        for col in columns:
            if col.split('_')[1] == 'MODAL A':
                insertion(T1, carga, col, 'MODAL A', col.split('_')[-1])
            elif col.split('_')[1] == 'MODAL B':
                insertion(T1, carga, col, 'MODAL B', col.split('_')[-1])
    elif 'SALDO' in i.upper():
        saldo_yes = True
        dia = i.upper().split(
            'SALDO')[-1].replace('.PDF', '').strip().replace('.', '/')
        T2 = transformation(result, T2, dia)
        T2 = T2.groupby(T2.index).sum()
if not T1.empty:
    excel_writer('T1', T1)
total_descarga = pd.DataFrame()
for i in FILES_PDF:
    descarga = pd.DataFrame(pdf_excel_reader(i))
    descarga.loc[3, :] = descarga.columns.map(lambda x: x.split('\t')[0])
    total_descarga = pd.concat(
        [total_descarga, descarga], axis=1, ignore_index=True)
if not total_descarga.empty:
    total_descarga = total_descarga.transpose()
    total_descarga[1] = total_descarga[1].map(
        lambda x: 'MODAL B' if 'TC' in x else 'MODAL A')
    # total_descarga[2] = total_descarga[2].map(lambda x: x.replace('.', '')).astype(int)
    somatorio = total_descarga.groupby([0, 1, 3]).sum().reset_index()
    somatorio['id'] = 'descarga_'+somatorio[1]+'_'+somatorio[3]
    for index in range(0, len(somatorio)):
        T2[somatorio.iloc[index]['id']][somatorio.iloc[index]
                                        [0]] = somatorio.iloc[index][2]
for i in FILES_EXCEL:
    if 'molde' in i:
        pass
    elif 'SEN' in i.upper():
        pass
    elif 'BASIS' in i.upper():
        total_R1_S1, ORGAO = excel_reader(i, registro)
        total_R1_S1 = pd.DataFrame.from_dict(total_R1_S1)
        total_R1_S1.set_index('dia', inplace=True)
        total_R1_S1.index.name = None
        if ORGAO == 'R1':
            R1 = pd.concat([R1, total_R1_S1],
                           axis=0, ignore_index=False)
        elif ORGAO == 'T2':
            for i in total_R1_S1.columns:
                total_R1_S1[i] = total_R1_S1[i].astype(int)
            T2 = pd.concat([T2, total_R1_S1],
                           axis=0, ignore_index=False)
            T2 = T2.groupby(T2.index).sum()
            if saldo_yes:
                T2['mercado_X1'] = T2['mercado_X1'] - \
                    T2['envio_T2_X1']
                T2['mercado_Y2'] = T2['mercado_Y2'] - \
                    T2['envio_T2_Y2']
                T2['mercado_Y1'] = T2['mercado_Y1'] - \
                    T2['envio_T2_Y1']
if not T2.empty:
    excel_writer('T2', T2)
if not registro.empty:
    excel_writer('registro', registro)
if not R1.empty:
    excel_writer('R1', R1)
for i in FILES_EXCEL:
    if 'SEN' in i.upper():
        total_saida = excel_reader(i, registro)
        total_R2 = pd.concat([total_saida['Unnamed: 3'][2:8].iloc[0:2],
                              total_saida['Unnamed: 3'][2:8].iloc[-1:],
                              total_saida['Unnamed: 5'][2:8].iloc[0:2],
                              total_saida['Unnamed: 5'][2:8].iloc[-1:],
                              total_saida['Unnamed: 9'][2:8].iloc[0:2],
                              total_saida['Unnamed: 9'][2:8].iloc[-1:]],
                             axis=0).to_list()
        dia = total_saida['Unnamed: 1'][14].strftime('%d/%m/%Y')
        columns = R2.columns.to_list()
        IND = 0
        total_R2_dict = {}
        for col in columns:
            total_R2_dict[col] = [total_R2[IND]]
            IND += 1
        total_R2_dict['dia'] = [dia]
        total_R2 = pd.DataFrame.from_dict(total_R2_dict)
        total_R2.set_index('dia', inplace=True)
        total_R2.index.name = None
        R2 = pd.concat([R2, total_R2], axis=0, ignore_index=False)
    elif 'ABERTURA' in i.upper():
        result, dia = S1_excel_reader(i)
        S1 = transformation(result, S1, dia)
        S1 = S1.groupby(S1.index).sum()
    elif 'DESCARGA' in i.upper():
        S1_excel_reader(i)

if not R2.empty:
    excel_writer('R2', R2)
if not S1.empty:
    excel_writer('S1', S1)
trash(FILES_EXCEL)
trash(FILES_PDF)
