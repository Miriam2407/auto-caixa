from tkinter import *
import tkinter as tk
from tkinter import ttk
from caixaRede import Rede
from caixaCielo import Cielo
from caixaGetnet import Getnet
import pandas as pd



interface = Tk()

interface.title("Automatização Cartões Caixa")
interface.geometry('400x400')

#label escolha de operadora
operdoralabel = ttk.Label(interface, text="Escolha uma operadora de cartão:")
operdoralabel.pack(pady=10)

#combobox para selecionar operadora
listOperadora = ["REDE", "CIELO", "GETNET"]

selectoperadora = tk.StringVar()
cb_operadora=ttk.Combobox(interface,textvariable=selectoperadora, values=listOperadora)
cb_operadora.set("GETNET")
cb_operadora.pack(pady=20)

#escolha a unidade
unidadelabel = ttk.Label(interface, text="Escolha uma unidade:")
unidadelabel.pack(pady=10)


#combobox para selecionar unidade
selectunidade = tk.StringVar()
listUnidade = ["Ibirapuera", "Vila Mariana", "ProFIV", "Campinas", "Cenafert", "Brasilia"]

cb_unidade=ttk.Combobox(interface,textvariable=selectunidade, values=listUnidade)
cb_unidade.set("Ibirapuera")
cb_unidade.pack(pady=20)        
 

#escolha de data para aquivo
datalabel = ttk.Label(interface, text="Escreva a data para o nome do aquivo (separando por pontos)")
datalabel.pack(pady=10)

#entry para data
selectdata = tk.StringVar()
campodata = ttk.Entry(interface, textvariable=selectdata)
campodata.pack(pady=10)


def selecionar_planilhas():
    operadora = selectoperadora.get()
    unidade = selectunidade.get()

    planilha1 = None
    planilha2 = None

    if operadora == "GETNET":
        if unidade == "Ibirapuera":
            planilha1 = 'data/notas_ibira.xlsx'
            planilha2 = 'data/vendas_ibira.xlsx'
        elif unidade == "Vila Mariana":
            planilha1 = 'data/notas_vila.xlsx'
            planilha2 = 'data/vendas_vila.xlsx'
        elif unidade == "ProFIV":
            planilha1 = 'data/notas_pel.xlsx'
            planilha2 = 'data/vendas_pel.xlsx'
        elif unidade == "Campinas":
            planilha1 = 'data/notas_campinas.xlsx'
            planilha2 = 'data/vendas_campinas.xlsx'
        elif unidade == "Cenafert":
            planilha1 = 'data/notas_cenafert.xlsx'
            planilha2 = 'data/vendas_cenafert.xlsx'
        else:
            planilha1 = 'data/notas_fiv.xlsx'
            planilha2 = 'data/vendas_fiv.xlsx'  

    elif operadora == "REDE":
        if unidade == "Ibirapuera":
            planilha1 = 'data/notas_ibira.xlsx'
            planilha2 = 'data/vendas_sa.csv'
        elif unidade == "Vila Mariana":
            planilha1 = 'data/notas_vila.xlsx'
            planilha2 = 'data/vendas_sa.csv'
        elif unidade == "ProFIV":
            planilha1 = 'data/notas_pel.xlsx'
            planilha2 = 'data/vendas_pel.csv'
        elif unidade == "Campinas":
            planilha1 = 'data/notas_campinas.xlsx'
            planilha2 = 'data/vendas_campinas.csv'
        elif unidade == "Cenafert":
            planilha1 = 'data/notas_cenafert.xlsx'
            planilha2 = 'data/vendas_cenafert.csv'
        else:
            planilha1 = 'data/notas_fiv.xlsx'
            planilha2 = 'data/vendas_fiv.csv'

    elif operadora == "CIELO":
        if unidade == "Ibirapuera":
            planilha1 = 'data/notas_ibira.xlsx'
            planilha2 = 'data/vendas_ibira.xlsx'
        elif unidade == "Vila Mariana":
            planilha1 = 'data/notas_vila.xlsx'
            planilha2 = 'data/vendas_vila.xlsx'
        elif unidade == "ProFIV":
            planilha1 = 'data/notas_pel.xlsx'
            planilha2 = 'data/vendas_pel.xlsx'
        elif unidade == "Campinas":
            planilha1 = 'data/notas_campinas.xlsx'
            planilha2 = 'data/vendas_campinas.xlsx'
        elif unidade == "Cenafert":
            planilha1 = 'data/notas_cenafert.xlsx'
            planilha2 = 'data/vendas_cenafert.xlsx'
        else:
            planilha1 = 'data/notas_fiv.xlsx'
            planilha2 = 'data/vendas_fiv.xlsx'    
    return planilha1, planilha2



    
def selecionar_operadora():
    operadora = selectoperadora.get()
    unidade = selectunidade.get()
    data = selectdata.get()

    if operadora == "REDE":
        def iniciarRede():
            planilha1, planilha2 = selecionar_planilhas()
            arquivoRede = Rede.iniciar(planilha1, planilha2, operadora, unidade, data)
            return arquivoRede
        iniciarRede()
    elif operadora == "CIELO":
        def inicarCielo():
            planilha1, planilha2 = selecionar_planilhas()
            arquivoCielo = Cielo.iniciar(planilha1, planilha2, operadora, unidade, data)
            return arquivoCielo
        inicarCielo()
    elif operadora == "GETNET":
        def inicarCielo():
            planilha1, planilha2 = selecionar_planilhas()
            arquivoCielo = Getnet.iniciar(planilha1, planilha2, operadora, unidade, data)
            return arquivoCielo
        inicarCielo()



#botão para criar arquivo
botaocriar = ttk.Button(interface, text="Criar arquivo", command=selecionar_operadora)
botaocriar.pack()

#resultado
resultadolabel = ttk.Label(interface, text='')
resultadolabel.pack(pady=10)

interface.mainloop()