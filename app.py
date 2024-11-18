"""GUI
"""

from tkinter import Button, Label, Toplevel, Text, END, Tk
from tkinter import messagebox
import pyperclip
import pandas as pd
from script_final import sap, excel_reader, pdf_reader, excel_write


def tips():
    """instructions
    """
    tips_janela = Toplevel(janela)
    tips_janela.title('Tips')
    tips_janela.minsize(width=590, height=380)
    tips_janela.config(padx=10, pady=10)
    dica = Text(tips_janela, height=45, width=68)
    dica.insert(END, "1)Evite usar os comandos SAP: eles são bots de clique\n\
para operar o SAP e obter as informações, mas funcionam\n\
com tempos de carregamento e localizações de programas\n\
específicos do computador de quem criou isso.")
    dica.insert(END, "\n\nObtenha as planilhas de Excel no SAP para Rio Verde(1120) e\n\
Rondo(1109) e salve na pasta desse programa com os nomes \n\
Planilha Basis (1)' e 'Planilha Basis (2)' respectivamente,\n\
esse já é o nome padrão para salvar. A planilha do SAP deve\n\
conter apenas um dia de operação.")
    dica.insert(END, "\n\nPara obter o valor total movimentado, você precisa seguir os\n\
seguintes passos no Excel com a planilha Basis aberta:")
    dica.insert(END, "\nInserir>Tabela Dinâmica>Configurações:\n\
Produto na linha;Nome no Filtro;Qtde Medida nos Valores\n\
Coloque o filtro na Basis(1) da tabela dinâmica como\n\
TERMINAIS RIO VERDE E SINOP caso esteja mexendo na\n\
'Planilha Basis(2)' Nesse momento a aba da Tabela \n\
dinâmica precisa ter como nome\n\
'Planilha2' e a da tabela principal 'Planilha1'.")
    dica.insert(END,
                '\n\n2)Baixe arquivo de Estoque e MOV BR do TCT na pasta \ndesse programa')
    dica.insert(END, '\n\n3)Baixe a planilha excel da movimentação de Senador \n\
Canedo na pasta desse programa.')
    dica.insert(END,
                '\n\n4)Baixe o saldo e os relatórios de descarga rodoviária\n\
e ferroviária do TECIAP na pasta desse programa.')
    dica.insert(END, "\n\n5)Mantenha a planilha 'registro.xlxs' na pasta desse\n\
programa a todo momento, lá é onde as movimentações e descargas são computadas.\n\
Após usar o programa basta abrir essa planilha e copiar\n\
os dados para as planilhas do colaborativo.")
    dica.insert(END, "\n\n6)PDF Estoque do TCT precisa ter a palavra 'Estoque' no nome.\n\
PDF Movimentação TCT precisa ter 'Mov' no nome.\n\
PDF Saldo TECIAP precisa ter 'Saldo' no nome.\n\
PDF Descarga rodoviária precisa ter 'Descarga' no nome.\n\
PDF Descarga ferroviária precisa ter 'Ferroviária' ou\n\
'Ferroviário' no nome.\n\
Excel de movimentações de Canedo precisa ter 'Senador'\n\
no nome.")
    dica.insert(END, '\n\n7)Esses são nomes padrões desses arquivos até o dia\n\
18/11/2024, então se não tiver mudado,\n\
não é preciso se atentar para isso.')
    dica.insert(END, '\n\n8)O mesmo vale para a formatação, esse programa é\n\
feito para funcionar com a formatação desses arquivos\n\
conforme estão até o dia 18/11/2024\n\
se houverem mudanças após esse dia pode ser que o\n\
programa não funcione mais.')
    dica.insert(END, '\n\n9)O programa não consegue ler PDF de múltiplos\n\
dias ao mesmo tempo, por isso separe os arquivos por dia e\n\
coloque-os na pasta do programa um dia por vez, rode o programa,\n\
tire os arquivos daquele dia e coloque os do próximo para rodar \n\
o programa novamente.')
    dica.insert(
        END, '\n\n10)Apague os arquivos dos terminais quando acabar\nde usar o programa.')
    dica.pack(expand=True)
    dica.configure(state="disabled")
    tips_janela.mainloop()


def sap_rv():
    """button calling function
    """
    nova('sap', 'RV')


def sap_sinop():
    """button calling function
    """
    nova('sap', 'SINOP')


def excel_rv():
    """button calling function
    """
    nova('excel', 'RV')


def excel_sinop():
    """button calling function
    """
    nova('excel', 'SINOP')


def pdf_rondo():
    """button calling function
    """
    nova('pdf', 'RONDO')


def nova(function, polo):
    """showing data
    """
    def copiar_rodo():
        pyperclip.copy(text_gasoa1.get('1.0', END).strip()+'\t' +
                       text_s500_1.get('1.0', END).strip()+'\t'+text_s10_1.get('1.0', END).strip())
        aviso1 = Label(nova_janela, text='Copiado!')
        aviso1.grid(column=2, row=3)

    def copiar_ferro():
        pyperclip.copy(text_gasoa2.get('1.0', END).strip()+'\t' +
                       text_s500_2.get('1.0', END).strip() + '\t' + text_s10_2.get('1.0', END).strip())
        aviso1 = Label(nova_janela, text='Copiado!')
        aviso1.grid(column=2, row=7)

    def copiar_rodo2():
        pyperclip.copy(text_gasoa3.get('1.0', END).strip()+'\t' +
                       text_s500_3.get('1.0', END).strip()+'\t'+text_s10_3.get('1.0', END).strip())
        aviso2 = Label(nova_janela, text='Copiado!')
        aviso2.grid(column=6, row=3)

    def copiar_ferro2():
        pyperclip.copy(text_gasoa4.get('1.0', END).strip()+'\t' +
                       text_s500_4.get('1.0', END).strip()+'\t' +
                       text_s10_4.get('1.0', END).strip())
        aviso2 = Label(nova_janela, text='Copiado!')
        aviso2.grid(column=6, row=7)

    def copiar_rodo3():
        # pylint: disable=used-before-assignment
        pyperclip.copy(text_gasoa5.get('1.0', END).strip()+'\t' +
                       text_s500_5.get('1.0', END).strip()+'\t'+text_s10_5.get('1.0', END).strip())
        aviso3 = Label(nova_janela, text='Copiado!')
        aviso3.grid(column=2, row=11)

    def copiar_ferro3():
        # pylint: disable=used-before-assignment
        pyperclip.copy(text_gasoa6.get('1.0', END).strip()+'\t' +
                       text_s500_6.get('1.0', END).strip()+'\t' +
                       text_s10_6.get('1.0', END).strip())
        aviso3 = Label(nova_janela, text='Copiado!')
        aviso3.grid(column=6, row=11)

    def mov():
        # pylint: disable=used-before-assignment
        mov_janela = Toplevel(nova_janela)
        mov_janela.title('Movimentação')
        mov_janela.config(padx=10, pady=10)

        movimentacao = Text(mov_janela, height=30, width=150)
        movimentacao.grid(column=0, row=0)
        if function == "excel" and polo == 'SINOP':
            movimentacao.insert(END, resultado_sinop[1])
            pergunta = messagebox.askyesno(
                "Movimentações", "Inserir movimentações no 'registro.xlxs'?")
            if pergunta:
                try:
                    excel_write(resultado_sinop[1], 'Movimentação')
                except Exception as e:  # pylint: disable=broad-except
                    messagebox.showerror(
                        "Error", f'{e},{e.args}')
        elif function == "excel" and polo == 'RV':
            movimentacao.insert(END, resultado_rv[1])
            pergunta = messagebox.askyesno(
                "Movimentações", "Inserir movimentações no 'registro.xlxs'?")
            if pergunta:
                try:
                    excel_write(resultado_rv[1], 'Movimentação')
                except Exception as e:  # pylint: disable=broad-except
                    messagebox.showerror(
                        "Error", f'{e},{e.args}')
        elif function == "pdf":
            movimentacao.insert(END, resultado_rondo[-2])
            pergunta = messagebox.askyesno(
                "Movimentações", "Inserir movimentações no 'registro.xlxs'?")
            if pergunta:
                try:
                    excel_write(resultado_rondo[-2], 'Movimentação')
                except Exception as e:  # pylint: disable=broad-except
                    messagebox.showerror(
                        "Error", f'{e},{e.args}')

    def descarga():
        descarga_janela = Toplevel(nova_janela)
        descarga_janela.title('Descarga')
        descarga_janela.config(padx=10, pady=10)

        descarga = Text(descarga_janela, height=30, width=110)
        descarga.grid(column=0, row=0)
        if function == "excel" and polo == 'RV':
            descarga.insert(END, resultado_rv[3].to_string(
                index=False, header=False).replace(' DSL S10', 'DSL S10').replace(
                    '   GASOA', 'GASOA'
            ))
            pergunta = messagebox.askyesno(
                "Descarregamento", "Inserir descarregamentos no 'registro.xlxs'?")
            if pergunta:
                try:
                    excel_write(resultado_rv[4], 'Descarga')
                except Exception as e:  # pylint: disable=broad-except
                    messagebox.showerror(
                        "Error", f'{e},{e.args}')
        elif function == 'pdf':
            descarga.insert(END, "\n".join(
                [i+'\t'+j for i, j in resultado_rondo[2].items()]))
            pergunta = messagebox.askyesno(
                "Descarregamento", "Inserir descarregamentos no 'registro.xlxs'?")
            if pergunta:
                try:
                    excel_write(resultado_rondo[2], 'Descarga')
                except Exception as e:  # pylint: disable=broad-except
                    messagebox.showerror(
                        "Error", f'{e},{e.args}')

    nova_janela = Toplevel(janela)
    nova_janela.title('Dados')
    nova_janela.config(padx=10, pady=10)

    main_label1 = Label(nova_janela, text='Rodoviário')
    main_label1.grid(column=1, row=0)
    main_label2 = Label(nova_janela, text='Ferroviário')
    main_label2.grid(column=1, row=4)
    main_label3 = Label(nova_janela, text='Rodoviário')
    main_label3.grid(column=5, row=0)
    main_label4 = Label(nova_janela, text='Ferroviário')
    main_label4.grid(column=5, row=4)
    main_label5 = Label(nova_janela, text='Saldo')
    main_label5.grid(column=1, row=13)
    main_label6 = Label(nova_janela, text='Saldo')
    main_label6.grid(column=5, row=13)

    gasoa_label1 = Label(nova_janela, text='GASOA')
    gasoa_label1.grid(column=0, row=1)
    gasoa_label2 = Label(nova_janela, text='GASOA')
    gasoa_label2.grid(column=0, row=5)
    gasoa_label3 = Label(nova_janela, text='GASOA')
    gasoa_label3.grid(column=4, row=1)
    gasoa_label4 = Label(nova_janela, text='GASOA')
    gasoa_label4.grid(column=4, row=5)

    s10_label1 = Label(nova_janela, text='DSL S10')
    s10_label1.grid(column=2, row=1)
    s10_label2 = Label(nova_janela, text='DSL S10')
    s10_label2.grid(column=2, row=5)
    s10_label3 = Label(nova_janela, text='DSL S10')
    s10_label3.grid(column=6, row=1)
    s10_label4 = Label(nova_janela, text='DSL S10')
    s10_label4.grid(column=6, row=5)

    s500_label1 = Label(nova_janela, text='DSL S500')
    s500_label1.grid(column=1, row=1)
    s500_label2 = Label(nova_janela, text='DSL S500')
    s500_label2.grid(column=1, row=5)
    s500_label3 = Label(nova_janela, text='DSL S500')
    s500_label3.grid(column=5, row=1)
    s500_label4 = Label(nova_janela, text='DSL S500')
    s500_label4.grid(column=5, row=5)

    text_gasoa1 = Text(nova_janela, height=2, width=7)
    text_gasoa1.grid(column=0, row=2)
    text_gasoa2 = Text(nova_janela, height=2, width=7)
    text_gasoa2.grid(column=0, row=6)
    text_gasoa3 = Text(nova_janela, height=2, width=7)
    text_gasoa3.grid(column=4, row=2)
    text_gasoa4 = Text(nova_janela, height=2, width=7)
    text_gasoa4.grid(column=4, row=6)

    text_s500_1 = Text(nova_janela, height=2, width=7)
    text_s500_1.grid(column=1, row=2)
    text_s500_2 = Text(nova_janela, height=2, width=7)
    text_s500_2.grid(column=1, row=6)
    text_s500_3 = Text(nova_janela, height=2, width=7)
    text_s500_3.grid(column=5, row=2)
    text_s500_4 = Text(nova_janela, height=2, width=7)
    text_s500_4.grid(column=5, row=6)

    text_s10_1 = Text(nova_janela, height=2, width=7)
    text_s10_1.grid(column=2, row=2)
    text_s10_2 = Text(nova_janela, height=2, width=7)
    text_s10_2.grid(column=2, row=6)
    text_s10_3 = Text(nova_janela, height=2, width=7)
    text_s10_3.grid(column=6, row=2)
    text_s10_4 = Text(nova_janela, height=2, width=7)
    text_s10_4.grid(column=6, row=6)

    button_copy1 = Button(
        nova_janela, text='Copiar Rodoviário', command=copiar_rodo)
    button_copy1.grid(column=1, row=3)
    button_copy2 = Button(
        nova_janela, text='Copiar Ferroviário', command=copiar_ferro)
    button_copy2 .grid(column=1, row=7)
    button_copy3 = Button(nova_janela, text='Copiar Rodoviário',
                          command=copiar_rodo2)
    button_copy3.grid(column=5, row=3)
    button_copy4 = Button(
        nova_janela, text='Copiar Ferroviário', command=copiar_ferro2)
    button_copy4 .grid(column=5, row=7)

    esp1 = Label(nova_janela, text='                      ')
    esp1.grid(column=3, row=0)
    esp2 = Label(nova_janela, text='                      ')
    esp2.grid(column=1, row=11)

    movs1 = Button(nova_janela, text='Movimentações', command=mov)
    movs1.grid(column=1, row=12)
    movs2 = Button(nova_janela, text='Descargas', command=descarga)
    movs2.grid(column=5, row=12)

    saldo1_gasoa = Label(nova_janela, text='0')
    saldo1_gasoa.grid(column=0, row=14)
    saldo1_s10 = Label(nova_janela, text='0')
    saldo1_s10.grid(column=2, row=14)
    saldo1_s500 = Label(nova_janela, text='0')
    saldo1_s500.grid(column=1, row=14)

    saldo2_gasoa = Label(nova_janela, text='0')
    saldo2_gasoa.grid(column=4, row=14)
    saldo2_s10 = Label(nova_janela, text='0')
    saldo2_s10.grid(column=6, row=14)
    saldo2_s500 = Label(nova_janela, text='0')
    saldo2_s500.grid(column=5, row=14)

    labels = [main_label1, gasoa_label1, s10_label1, s500_label1, s10_label3,
              s500_label3, gasoa_label3, main_label3, gasoa_label4,
              s10_label4,  s500_label4, main_label4,
              s10_label2, s500_label2, main_label2, gasoa_label2]
    text = [text_gasoa1, text_s10_1, text_s500_1, text_gasoa3, text_s10_3, text_s500_3,
            text_gasoa4, text_s10_4,
            text_s500_4, text_gasoa2, text_s10_2, text_s500_2]
    buttons = [button_copy1, button_copy3,
               movs1, movs2, button_copy4, button_copy2]
    saldo = [main_label5, saldo1_gasoa, saldo1_s10, saldo1_s500,
             main_label6, saldo2_gasoa, saldo2_s10, saldo2_s500]
    if function == 'excel':
        if polo == 'SINOP':
            resultado_sinop = excel_reader('SINOP')
            if resultado_sinop == 'Sem entradas':
                nova_janela.destroy()
                messagebox.showerror(
                    "Error", "Sem arquivos de SINOP - Planilha Basis(2)")
            else:
                entrada_sinop = resultado_sinop[0]
                main_label1.config(text='Carregamento para SINOP')
                for i, j in entrada_sinop.items():
                    if i == 'PB.620':
                        text_gasoa1.insert(END, j)
                    elif i == 'PB.658':
                        text_s500_1.insert(END, j)
                    elif i == 'PB.6DH':
                        text_s10_1.insert(END, j)
                text_s10_1.configure(state="disabled")
                text_gasoa1.configure(state="disabled")
                text_s500_1.configure(state="disabled")
                for i in labels:
                    if (i is not main_label1) and (i is not gasoa_label1) and (i is not s10_label1) and\
                            (i is not s500_label1):
                        i.grid_forget()
                for i in buttons:
                    if (i is not button_copy1) and (i is not movs1):
                        i.grid_forget()
                for i in saldo:
                    i.grid_forget()
                for i in text:
                    if (i is not text_gasoa1) and (i is not text_s10_1) and (i is not text_s500_1):
                        i.grid_forget()

        elif polo == 'RV':
            resultado_rv = excel_reader('RV')
            if resultado_rv == 'Sem dados':
                nova_janela.destroy()
                messagebox.showerror(
                    "Error", "Sem arquivos de Rio Verde")
            else:
                entrada_rv = resultado_rv[0]
                main_label1.config(text='Carregamento Senador Canedo')
                main_label3.config(text='Entrada - Descarga DTC')
                main_label4.config(text='Saída - Carregamento DTC')
                button_copy4.config(text='Copiar Rodoviário')
                for i, j in entrada_rv.items():
                    if i == 'PB.620':
                        text_gasoa1.insert(END, j)
                    elif i == 'PB.658':
                        text_s500_1.insert(END, j)
                    elif i == 'PB.6DH':
                        text_s10_1.insert(END, j)
                text_s10_1.configure(state="disabled")
                text_gasoa1.configure(state="disabled")
                text_s500_1.configure(state="disabled")

                entradas_dtc = pd.concat([resultado_rv[2]['Unnamed: 3'][2:8].iloc[0:2],
                                          resultado_rv[2]['Unnamed: 3'][2:8].iloc[-1:]],
                                         axis=0).to_list()
                saidas_dtc = pd.concat([resultado_rv[2]['Unnamed: 5'][2:8].iloc[0:2],
                                        resultado_rv[2]['Unnamed: 5'][2:8].iloc[-1:]],
                                       axis=0).to_list()
                saldos_dtc = pd.concat([resultado_rv[2]['Unnamed: 9'][2:8].iloc[0:2],
                                        resultado_rv[2]['Unnamed: 9'][2:8].iloc[-1:]],
                                       axis=0).to_list()

                text_gasoa3.insert(END, entradas_dtc[0])
                text_s500_3.insert(END, entradas_dtc[1])
                text_s10_3.insert(END, entradas_dtc[2])
                text_gasoa4.insert(END, saidas_dtc[0])
                text_s500_4.insert(END, saidas_dtc[1])
                text_s10_4.insert(END, saidas_dtc[2])
                new_saldos_dtc = []
                for i in saldos_dtc:
                    if len(str(i)) > 6:
                        new_saldos_dtc.append(
                            str(i)[0]+'.'+str(i)[1:4]+'.'+str(i)[-3:])
                    elif len(str(i)) < 7 and len(str(i)) > 3:
                        new_saldos_dtc.append(str(i)[0]+'.'+str(i)[1:])
                    else:
                        new_saldos_dtc.append(str(i))
                saldo2_gasoa.config(text=new_saldos_dtc[0])
                saldo2_s500.config(text=new_saldos_dtc[1])
                saldo2_s10.config(text=new_saldos_dtc[2])
                text_s500_3.configure(state="disabled")
                text_s10_3.configure(state="disabled")
                text_gasoa3.configure(state="disabled")
                text_gasoa4.configure(state="disabled")
                text_s10_4.configure(state="disabled")
                text_s500_4.configure(state="disabled")
                for i in labels[12:]:
                    i.grid_forget()
                for i in buttons[5:]:
                    i.grid_forget()
                for i in saldo:
                    if i is main_label5 or i is saldo1_gasoa or i is saldo1_s10 or i is saldo1_s500:
                        i.grid_forget()
                for i in text[9:]:
                    i.grid_forget()
    elif function == 'pdf':
        resultado_rondo = pdf_reader('RONDO')

        entrada_rondo = resultado_rondo[0]
        saida_rondo = resultado_rondo[-1]
        total_teciap = resultado_rondo[1]
        divisao_modal = resultado_rondo[-3]

        valores_teciap = [
            '%.3f' % n for n in total_teciap[total_teciap.columns[2]].to_list()]
        for i in range(0, len(valores_teciap)):
            if len(valores_teciap[i]) > 7:
                valores_teciap[i] = valores_teciap[i][0] + \
                    '.'+valores_teciap[i][1:]
        mercado_teciap = [
            '%.3f' % n for n in total_teciap[total_teciap.columns[1]].to_list()]
        mercado_teciap = [i.replace('0.000', '0') for i in mercado_teciap]
        main_label1.config(text='Entrada TCT')
        main_label2.config(text='Saída TCT Ferroviário')
        main_label3.config(text='Entrada TECIAP - Rodoviário')
        main_label4.config(text='Entrada TECIAP - Ferroviário')

        saldo2_gasoa.config(text=valores_teciap[0])
        saldo2_s500.config(text=valores_teciap[1])
        saldo2_s10.config(text=valores_teciap[2])

        main_label7 = Label(nova_janela, text='Saída TCT Rodoviário')
        main_label7.grid(column=1, row=8)
        main_label8 = Label(nova_janela, text='Mercado TECIAP')
        main_label8.grid(column=5, row=8)

        gasoa_label5 = Label(nova_janela, text='GASOA')
        gasoa_label5.grid(column=0, row=9)
        gasoa_label6 = Label(nova_janela, text='GASOA')
        gasoa_label6.grid(column=4, row=9)

        s10_label5 = Label(nova_janela, text='DSL S10')
        s10_label5.grid(column=2, row=9)
        s10_label6 = Label(nova_janela, text='DSL S10')
        s10_label6.grid(column=6, row=9)

        s500_label5 = Label(nova_janela, text='DSL S500')
        s500_label5.grid(column=1, row=9)
        s500_label6 = Label(nova_janela, text='DSL S500')
        s500_label6.grid(column=5, row=9)

        text_gasoa5 = Text(nova_janela, height=2, width=7)
        text_gasoa5.grid(column=0, row=10)
        text_gasoa6 = Text(nova_janela, height=2, width=7)
        text_gasoa6.grid(column=4, row=10)

        text_s10_5 = Text(nova_janela, height=2, width=7)
        text_s10_5.grid(column=2, row=10)
        text_s10_6 = Text(nova_janela, height=2, width=7)
        text_s10_6.grid(column=6, row=10)

        text_s500_5 = Text(nova_janela, height=2, width=7)
        text_s500_5.grid(column=1, row=10)
        text_s500_6 = Text(nova_janela, height=2, width=7)
        text_s500_6.grid(column=5, row=10)

        button_copy5 = Button(
            nova_janela, text='Copiar Rodoviário', command=copiar_rodo3)
        button_copy5.grid(column=1, row=11)
        button_copy6 = Button(
            nova_janela, text='Copiar Mercado', command=copiar_ferro3)
        button_copy6.grid(column=5, row=11)
        try:
            text_gasoa1.insert(END, entrada_rondo['Entradas'][2])
        except TypeError:
            text_gasoa1.insert(END, 0)
        try:
            text_gasoa2.insert(END, saida_rondo.iloc[0][:][0])
            text_gasoa5.insert(END, saida_rondo.iloc[1][:][0])
        except AttributeError:
            text_gasoa2.insert(END, 0)
            text_gasoa5.insert(END, 0)

        text_gasoa6.insert(END, mercado_teciap[0])

        try:
            text_s500_1.insert(END, entrada_rondo['Entradas'][8])
        except TypeError:
            text_s500_1.insert(END, 0)

        try:
            text_s500_2.insert(END, saida_rondo.iloc[0][:][1])
            text_s500_5.insert(END, saida_rondo.iloc[1][:][1])
        except AttributeError:
            text_s500_2.insert(END, 0)
            text_s500_5.insert(END, 0)

        text_s500_6.insert(END, mercado_teciap[1])

        try:
            text_s10_1.insert(END, entrada_rondo['Entradas'][12])
        except TypeError:
            text_s10_1.insert(END, 0)
        try:
            text_s10_2.insert(END, saida_rondo.iloc[0][:][2])
            text_s10_5.insert(END, saida_rondo.iloc[1][:][2])
        except AttributeError:
            text_s10_2.insert(END, 0)
            text_s10_5.insert(END, 0)
        text_s10_6.insert(END, mercado_teciap[2])

        # pylint: disable=no-member
        try:
            saldo1_gasoa.config(
                text=entrada_rondo[entrada_rondo.columns[-2]][2])
            saldo1_s10.config(
                text=entrada_rondo[entrada_rondo.columns[-2]][12])
            saldo1_s500.config(
                text=entrada_rondo[entrada_rondo.columns[-2]][8])
        except AttributeError:
            saldo1_gasoa.config(text='Sem arquivo TCT')
            saldo1_s10.config(text='Sem arquivo TCT')
            saldo1_s500.config(text='Sem arquivo TCT')

        entrada_teciap = [str(i).replace('0.0', '0')
                          for i in total_teciap[total_teciap.columns[0]]]

        if divisao_modal == '':
            text_gasoa3.insert(
                END, entrada_teciap[0])
            text_s500_3.insert(
                END, entrada_teciap[1])
            text_s10_3.insert(
                END, entrada_teciap[2])
        else:
            text_gasoa4.insert(
                END, str(format(divisao_modal['fe_GASOA'], '.3f')).replace('0.000', '0'))
            text_s500_4.insert(
                END, str(format(divisao_modal['fe_DSL500'], '.3f')).replace('0.000', '0'))
            text_s10_4.insert(
                END, str(format(divisao_modal['fe_DSL10'], '.3f')).replace('0.000', '0'))

            text_gasoa3.insert(
                END, str(format(divisao_modal['ro_GASOA'], '.3f')).replace('0.000', '0'))
            text_s500_3.insert(
                END, str(format(divisao_modal['ro_DSL500'], '.3f')).replace('0.000', '0'))
            text_s10_3.insert(
                END, str(format(divisao_modal['ro_DSL10'], '.3f')).replace('0.000', '0'))

        text_gasoa1.configure(state="disabled")
        text_gasoa2.configure(state="disabled")
        text_gasoa3.configure(state="disabled")
        text_gasoa4.configure(state="disabled")
        text_gasoa5.configure(state="disabled")

        text_s500_1.configure(state="disabled")
        text_s500_2.configure(state="disabled")
        text_s500_3.configure(state="disabled")
        text_s500_4.configure(state="disabled")
        text_s500_5.configure(state="disabled")

        text_s10_1.configure(state="disabled")
        text_s10_2.configure(state="disabled")
        text_s10_3.configure(state="disabled")
        text_s10_4.configure(state="disabled")
        text_s10_5.configure(state="disabled")
    elif function == 'sap':
        if polo == 'SINOP':
            nova_janela.destroy()
            sap('1109')
        elif polo == 'RV':
            nova_janela.destroy()
            sap('1120')

    nova_janela.mainloop()


janela = Tk()
janela.title('Scraping')
janela.minsize(width=270, height=450)
janela.config(padx=20, pady=20)
botao_op1 = Button(text='SAP Rio Verde', command=sap_rv)
botao_op1.place(x=20, y=10)
botao_op2 = Button(text='SAP SINOP', command=sap_sinop)
botao_op2.place(x=20, y=60)
botao_op3 = Button(text='Leitura Rio Verde', command=excel_rv)
botao_op3.place(x=20, y=110)
botao_op4 = Button(text='Leitura SINOP', command=excel_sinop)
botao_op4.place(x=20, y=160)
botao_op5 = Button(text='Leitura Rondonópolis', command=pdf_rondo)
botao_op5.place(x=20, y=210)
botao_op6 = Button(text='SAIR', command=exit)
botao_op6.place(x=20, y=300)
botao_op7 = Button(text='Instruções', command=tips)
botao_op7.place(x=20, y=350)

janela.mainloop()
