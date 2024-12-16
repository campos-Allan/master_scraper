"""variables creation
"""
import pandas as pd

registro = pd.DataFrame(data=None, columns=[
    'data_carga',
    'polo_1',
    'polo_2',
    'tipo',
    'id_1',
    'prod',
    'vol_carga',
    'eta',
    'data_descarga',
    'id_2',
    'vol_des',
    'obs'])
T1 = pd.DataFrame(data=None, columns=[
    'entrada_X1',
    'entrada_Y2',
    'entrada_Y1',
    'saida_MODAL A_X1',
    'saida_MODAL A_Y2',
    'saida_MODAL A_Y1',
    'saida_MODAL B_X1',
    'saida_MODAL B_Y2',
    'saida_MODAL B_Y1',
    'saldo_X1',
    'saldo_Y2',
    'saldo_Y1'])
T2 = pd.DataFrame(data=None, columns=[
    'descarga_MODAL A_X1',
    'descarga_MODAL A_Y2',
    'descarga_MODAL A_Y1',
    'descarga_MODAL B_X1',
    'descarga_MODAL B_Y2',
    'descarga_MODAL B_Y1',
    'mercado_X1',
    'mercado_Y2',
    'mercado_Y1',
    'envio_T2_X1',
    'envio_T2_Y2',
    'envio_T2_Y1',
    'saldo_X1',
    'saldo_Y2',
    'saldo_Y1'])
S1 = pd.DataFrame(data=None, columns=[
    'descarga_X1',
    'descarga_Y2',
    'descarga_Y1',
    'mercado_X1',
    'mercado_Y2',
    'mercado_Y1',
    'saldo_X1',
    'saldo_Y2',
    'saldo_Y1'])
R1 = pd.DataFrame(data=None, columns=[
    'saida_X1',
    'saida_Y2',
    'saida_Y1'])
R2 = pd.DataFrame(data=None, columns=[
    'descarga_X1',
    'descarga_Y2',
    'descarga_Y1',
    'mercado_X1',
    'mercado_Y2',
    'mercado_Y1',
    'saldo_X1',
    'saldo_Y2',
    'saldo_Y1'])
