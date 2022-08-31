import tkinter as tk
from tkinter import ttk
import numpy as np
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime


requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = [moeda for moeda in dicionario_moedas]

def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    data_cotacao_dividida= data_cotacao
    data_cotacao_dividida = data_cotacao_dividida.split('/')
    ano = data_cotacao_dividida[2]
    mes = data_cotacao_dividida[1]
    dia = data_cotacao_dividida[0]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    variavel_cotacao =requisicao_moeda.json()
    try:
        variavel_cotacao['status'] == 404
        label_textocotacao['foreground'] = 'red'
        label_textocotacao['text'] = f'A moeda "{moeda.upper()}" não existe. Insira uma moeda válida.'

    except:
        if len(variavel_cotacao) == 0:
            label_textocotacao['foreground'] = 'red'
            label_textocotacao['text'] = f'Não tem cotação da moeda "{moeda}" no dia {data_cotacao}. Por favor colocar um dia válido.'
        else:
            valor_moeda = float(variavel_cotacao[0]['bid'])
            label_textocotacao['foreground'] = 'black'
            label_textocotacao['text'] = f'A cotação da moeda "{moeda}" no dia {data_cotacao} foi de R${valor_moeda:.2f}'



def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'


def atualizar_cotacoes():
    df = pd.read_excel(var_caminho_arquivo.get())
    moedas = df.iloc[:,0]
    
    data_inicial = calendario_datainicial.get()
    ano_inicial = data_inicial[-4:]
    mes_inicial = data_inicial[3:5]
    dia_inicial = data_inicial[:2]
    
    data_final = calendario_datafinal.get()
    ano_final = data_final[-4:]
    mes_final = data_final[3:5]
    dia_final = data_final[:2]

    for moeda in moedas:
        link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/200?' \
               f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&' \
               f'end_date={ano_final}{mes_final}{dia_final}'
        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao['bid'])
            data = datetime.fromtimestamp(timestamp)
            data = data.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan
            df.loc[df.iloc[:,0] == moeda, data] = bid
    df.to_excel('Teste.xlsx', index=False)
    label_atualizarcotacoes['text'] = 'Arquivo Atualizado com Sucesso'


janela = tk.Tk()

ano_atual = int(str(datetime.today()).split('-')[0])

janela.title('Ferramenta de Cotação de Moedas')

label_cotacaomoeda = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid', background='#100646', foreground='white')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='NWSE', columnspan=3)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='NWSE')

label_selecionarmoeda = tk.Label(text='Selecionar moeda desejada', anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='NWSE', columnspan=2)

label_selecionardia = tk.Label(text='Selecionar o dia que deseja pegar a cotação', anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='NWSE', columnspan=2)

calendario_moeda = DateEntry(year=ano_atual, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='NWSE')

botao_pegarcotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, columnspan=2, padx=10, pady=10, sticky='NWSE')

# Cotação de Várias Moedas

label_cotacaovariasmoeda = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid', background='#100646', foreground='white')
label_cotacaovariasmoeda.grid(row=4, column=0, padx=10, pady=10, sticky='NWSE', columnspan=3)

label_selecionarquivo = tk.Label(text='Selecione um arquivo em excel com as Moedas na Coluna A')
label_selecionarquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')

#uma string armazenada dentro do TKinter, podemos utilizar ela nas funções
var_caminho_arquivo = tk.StringVar()

botao_selecionaraquivo = tk.Button(text='Selecionar Arquivo', command=selecionar_arquivo)
botao_selecionaraquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

label_arquivoselecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nswe')

label_datainicial = tk.Label(text='Data Inicial', anchor='e')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')

label_datafinal = tk.Label(text='Data Final', anchor='e')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

calendario_datainicial = DateEntry(year=ano_atual, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')

calendario_datafinal = DateEntry(year=ano_atual, locale='pt_br')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')

botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='NWSE')

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='NWSE')

janela.mainloop()
