"""graphical user interface
"""
import os
from tkinter import Button, Tk

def reg():
    os.system('main.py')
def readme():
    os.system("start " + 'readme.txt')
janela = Tk()
janela.title('Scraping')
janela.minsize(width=150, height=170)
janela.config(padx=20, pady=20)
botao_op1 = Button(text='Registrar', command=reg)
botao_op1.place(x=20, y=0)
botao_op2 = Button(text='Instruções', command=readme)
botao_op2.place(x=20, y=50)
botao_op3 = Button(text='SAIR', command=exit)
botao_op3.place(x=20, y=100)

janela.mainloop()