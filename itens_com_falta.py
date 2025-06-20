"""
Sistema de Consulta e Relatórios de Pedidos Não Liberados com Estoque Zerado
----------------------------------------------------------------------------

Este sistema automatiza o acesso ao sistema interno de uma empresa para coletar dados
de pedidos com produtos sem estoque disponível, utilizando Selenium para navegação,
Tkinter para interface gráfica e Pandas para manipulação de dados.

Funcionalidades:
- Acesso automático à intranet com login e senha via variáveis de ambiente.
- Navegação até a tela de posição de estoque de forma automatizada.
- Coleta de dados dos produtos sem estoque e seus pedidos relacionados.
- Geração de relatórios em planilhas Excel com backup automático.
- Interface gráfica com opções de seleção da unidade de distribuição.

---

System for Consulting and Reporting Unreleased Orders with Zero Stock
---------------------------------------------------------------------

This system automates access to a company intranet to collect data on orders
with products that are out of stock. It uses Selenium for browser automation,
Tkinter for GUI and Pandas for data handling.

Features:
- Automatic intranet login using credentials from environment variables.
- Automated navigation to the stock position screen.
- Collection of out-of-stock products and related order information.
- Excel report generation with automatic backup.
- GUI with distributor selection.

"""

from flask import Flask
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from time import sleep
from datetime import datetime
from tkinter import *
from tkinter import ttk, messagebox
from dotenv import load_dotenv
import os
import threading
from pathlib import Path
import logging
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
logging.debug('Iniciando a aplicacao')
# logging.disable(logging.CRITICAL)  # comente para habilitar
# logging.disable(logging.WARNING)  # comente para habilitar
# logging.disable(logging.INFO)  # comente para habilitar
# logging.disable(logging.DEBUG)  # comente para habilitar


load_dotenv()  # Carrega as variáveis do ..env
senha = os.getenv("senha")
login = os.getenv('login')
URL = os.getenv('URL')

pd.set_option('display.max_rows', None)  # Exibe todas as linhas
pd.set_option('display.width', None)  # Não limita a largura do terminal
pd.set_option('display.max_colwidth', None)  # Não limita a largura das colunas
pd.set_option('display.max_columns', None)  # exibe todas as colunas


app = Flask(__name__)
options = Options()
options.page_load_strategy = 'normal'
navegador = webdriver.Chrome(options=options)
navegador.minimize_window()
options.add_argument("--headless=new")

lista_geral = []
lista_saldo = []
lista_saldo_temp = []

# logging.disable(logging.info)

def criar_diretorio():
    logging.debug('Função criar_diretorio')
    caminho_0 = Path(r'C:\relato')
    caminho_0.mkdir(parents=True, exist_ok=True) # cria o diretorio, se ele nao existir
    logging.info('Cria o diretorio relato, se ele nao existir')

    caminho_1 = Path(r'C:\relato\cod_sem_estoque')
    caminho_1.mkdir(parents=True, exist_ok=True) # cria o diretorio, se ele nao existir
    logging.info('Cria o diretorio cod_sem_estoque, se ele nao existir')

    caminho_2 = Path(r'C:\relato\cod_sem_estoque\backup')
    caminho_2.mkdir(parents=True, exist_ok=True) # cria o diretorio, se ele nao existir
    logging.info('Cria o diretorio backup, se ele nao existir')


def acessa_intranet(dist):
    logging.debug('Função acessa_intranet')
    navegador.get(URL)
    navegador.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/input').clear()
    navegador.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/input').send_keys(
        login)

    sleep(2)
    navegador.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/input').send_keys(
        senha, Keys.ENTER)
    sleep(3)

    # Acessa a tela de pedidos
    navegador.find_element('xpath', '/html/body/form[1]/table/tbody/tr[2]/td[2]/select').click()
    navegador.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td[2]/select/option[7]').click()
    navegador.find_element('xpath', '//*[@id="navigation"]/ul/li[2]/a').click()  # clica em logistica
    navegador.find_element('xpath', '//*[@id="navigation"]/ul/li[2]/ul/a[4]').click()  # clica em posicao estoque

    #  clica em distribuidor
    navegador.find_element('xpath', '/html/body/form[2]/table[1]/tbody/tr[2]/td[1]/select').click()
    if dist == '1':
        # seleciona extrema
        navegador.find_element('xpath', '/html/body/form[2]/table[1]/tbody/tr[2]/td[1]/select/option[3]').click()
    else:
        # seleciona alhandra
        navegador.find_element('xpath', '/html/body/form[2]/table[1]/tbody/tr[2]/td[1]/select/option[2]').click()

    # processa
    navegador.find_element('xpath', '/html/body/form[2]/table[2]/tbody/tr/td[1]/input').click()
    logging.info('Abrindo tabela de códigos sem estoque')
    messagebox.showinfo('Processo Iniciado', 'Clique após o carregamento das informações')

def lista_codigos():
    logging.debug('\nFunçao lista_codigo')
    # teste clique
    navegador.find_element('xpath', '/html/body/form[2]/table[3]/tbody/tr[1]/th[1]').click()
    tr = 2
    while True:
        try:
            codigo = navegador.find_element('xpath', f'/html/body/form[2]/table[3]/tbody/tr[{tr}]/td[3]').text
            saldo = navegador.find_element('xpath', f'/html/body/form[2]/table[3]/tbody/tr[{tr}]/td[9]').text
            saldo = saldo.replace(' ', '')
            saldo = saldo.replace('.', '')
            logging.debug(f'\nAnálise linha {tr} | Código: {codigo} | Saldo: {saldo} |')
            navegador.find_element('xpath', f'/html/body/form[2]/table[3]/tbody/tr[{tr}]/td[1]/a/img').click()
            navegador.implicitly_wait(60)
            lista_interna(codigo, saldo)
            navegador.implicitly_wait(60)
            navegador.find_element('xpath', f'/html/body/form[2]/table[3]/tbody/tr[{tr}]/td[1]/a/img').click()
            tr += 2

        except:
            logging.warning('\n(Except) Fim da lista')
            lista_geral.extend(lista_interna(codigo, saldo))
            for i in lista_geral:
                logging.info(f'{i}')
            break

def lista_interna(codigo,saldo):
    logging.debug('\nFunçao lista_interna')
    tr = 3
    while True:
        navegador.implicitly_wait(10)
        print('█', end='')
        try:
            pedido = navegador.find_element('xpath', f'//*[@id="cpo{codigo}"]/table/tbody/tr/td/table/tbody/tr[{tr}]/td[1]').text
            vendedor = navegador.find_element('xpath', f'//*[@id="cpo{codigo}"]/table/tbody/tr/td/table[1]/tbody/tr[{tr}]/td[3]').text
            dta_emi = navegador.find_element('xpath', f'//*[@id="cpo{codigo}"]/table/tbody/tr/td/table[1]/tbody/tr[{tr}]/td[4]').text
            nome_fant = navegador.find_element('xpath', f'//*[@id="cpo{codigo}"]/table/tbody/tr/td/table[1]/tbody/tr[{tr}]/td[7]').text
            qt_n_liberado = navegador.find_element('xpath', f'//*[@id="cpo{codigo}"]/table/tbody/tr/td/table[1]/tbody/tr[{tr}]/td[10]').text
            qt_n_liberado = qt_n_liberado.replace('.', '')
            qt_n_liberado = qt_n_liberado.replace(' ', '')
            qt_n_liberado = float(qt_n_liberado)
            qt_n_liberado = int(qt_n_liberado)
            lista_temp = (codigo, saldo, pedido, vendedor, dta_emi, nome_fant, qt_n_liberado)
            lista_geral.append(lista_temp[:])
            tr += 1

        except:
            logging.warning('\n(Except) Fim da lista')
            break
    return lista_geral

def relatorio_pedidos():
    logging.debug('função relatorio_pedidos')
    while True:
        navegador.implicitly_wait(10)
        print('█', end='')
        try:
            pass
        except:
            pass

def converte_dataframe(lista_geral, dist):
    logging.debug('\nFunção converte_dataframe')
    dist = dist
    logging.debug(f'Dist: {dist}')
    df = pd.DataFrame(lista_geral, columns=['codigo', 'saldo', 'pedido', 'vendedor', 'dta_emi', 'nome fantasia', 'Quantidade'])

    df['Quantidade'] = pd.to_numeric(df['Quantidade'])  # Convertendo a coluna 'Quantidade' para inteiro
    df['saldo'] = pd.to_numeric(df['saldo'])  # Convertendo a coluna 'Quantidade' para inteiro

    # agrupando pelo codigo e pedido, e somando as quantidades
    df_resumido = df.groupby(['codigo', 'pedido'], as_index=False)['Quantidade'].sum()

    # pivotando o dataframe para que cada pedido seja uma coluna
    df_pivot = df_resumido.pivot(index='pedido', columns='codigo', values='Quantidade').fillna(0).astype(int)

    df_auxiliar = pd.DataFrame(df, columns=['codigo', 'saldo'])

    return df_pivot, df_auxiliar, dist


def salva_arquivo(df_pivot, df_auxiliar, dist):
    logging.debug('função salva_arquivo')
    # Obtém a data e hora atual no formato desejado (ano-mês-dia_hora-minuto-segundo)
    # data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H_%M')
    dist = 'Alhandra' if dist == '2' else 'Extrema'
    excel_path_backup = f'C:/relato/cod_sem_estoque/backup/codigos_sem_estoque_{dist}_{data_hora_atual}.xlsx'
    excel_path = f'C:/relato/cod_sem_estoque/codigos_sem_estoque_{dist}.xlsx'

    # Salva o DataFrame como Excel
    df_pivot.to_excel(excel_path, index=True)
    df_pivot.to_excel(excel_path_backup, index=True)
    lista_geral.clear()

    df_auxiliar.to_excel(f'C:/relato/cod_sem_estoque/lista_saldo_itens_indisponiveis_{dist}.xlsx', index=False)



class Frame:
    def __init__(self, master,  opcao_radio_button, botao_confirma):  # todas as variáveis aqui também devem estar na função main
        master.title('Pedidos Não Liberados')  # titulo do frame principal
        master.resizable(True, True)  # define se as dimensoes do frame podem ser alteradas

        style = ttk.Style()
        style.theme_use('alt')  # Usa um tema moderno
        style.configure('TLabelFrame',  borderwidth=0)
        style.configure('TButton', background='#0078D7', foreground='white')
        style.configure('TProgressbar', thickness=10)

        #  objetos do frame_um
        self.opcao_radio_button = opcao_radio_button
        self.botao_confirma = botao_confirma
        # configuração do frame
        self.label_frame_um = ttk.LabelFrame(master)  # frame (borda) 1 instanciando frame 'master'
        self.label_frame_um.config(relief=GROOVE,  text='Distribuidora', padding=(20, 10))
        self.label_frame_um.config(width=500, height=300)  # define o tamanho
        self.label_frame_um.pack(fill=BOTH, expand=False)
        self.opcao_radio_button = StringVar()
        self.botao_confirma = StringVar()

        # configuração dos objetos dentro do frame 1
        self.radio_opcao_1 = ttk.Radiobutton(self.label_frame_um)  # instanciando o label_frame_um
        self.radio_opcao_1.config(text='Extrema', variable=self.opcao_radio_button, value=1, command=self.retorno_radio_button)
        self.radio_opcao_1.grid(row=0, column=0, padx=5, pady=5, sticky=W)  # posiciona o objeto no frame

        self.radio_opcao_2 = ttk.Radiobutton(self.label_frame_um)  # instanciando o label_frame_um
        self.radio_opcao_2.config(text='Alhandra', variable=self.opcao_radio_button, value=2, command=self.retorno_radio_button)
        self.radio_opcao_2.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        self.botao_confirma = ttk.Button(self.label_frame_um)
        self.botao_confirma.config(text='Confirma', command=self.iniciar_thread)
        self.botao_confirma.grid(row=1, column=0, padx=5, pady=5, sticky=W)

        self.progress = ttk.Progressbar(self.label_frame_um, orient=HORIZONTAL, length=300, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=2, pady=10)


        #  objetos do frame_dois
        self.label_frame_dois = ttk.LabelFrame(master)  # frame (borda) 1 instanciando frame 'master'

        ttk.Label(self.label_frame_dois, text='Desenvolvido por Jean Lino Versão 4.1', font=('arial', 9, "italic")).grid(
            row=4, column=0, sticky='sw')

        self.label_frame_dois.pack(fill=BOTH, expand=False)

    def retorno_radio_button(self):
        logging.debug('def retorno_radio_button')
        a =self.opcao_radio_button.get()  # instancia o objeto, que retornará  valor do objeto
        if a == '1':
            logging.info('opção_1 selecionada')
        if a == '2':
            logging.info('opção_2 selecionada')
        return a

    def iniciar_thread(self):
        logging.debug('iniciar_thread')
        threading.Thread(target=self.retorno_botao_confirma, daemon=True).start()

    def retorno_botao_confirma(self):
        logging.debug('def retorno_botao_confirma')

        dist = self.retorno_radio_button()
        if dist != '1' and dist != '2':
            messagebox.showerror('Erro', 'Selecione uma opção')
        else:
            try:
                self.progress['value'] = 15
                self.label_frame_um.update_idletasks()
                criar_diretorio()

                self.progress['value'] = 30
                self.label_frame_um.update_idletasks()

                acessa_intranet(dist)
                self.label_frame_um.update_idletasks()
                self.progress['value'] = 45

                lista_codigos()
                self.progress['value'] = 60
                self.label_frame_um.update_idletasks()

                df_pivot, df_resumido, dist = converte_dataframe(lista_geral, dist)
                salva_arquivo(df_pivot, df_resumido, dist)
                self.progress['value'] = 75
                self.label_frame_um.update_idletasks()

                self.progress['value'] = 100
                navegador.close()
                messagebox.showinfo('Processo Concluído', 'Arquivo salvo em: C:\\relato\\cod_sem_estoque')
            except Exception as e:
                logging.error(e)
                messagebox.showerror('Erro', f'Ocorreu um erro: {str(e)}')


def main():
    logging.debug('def main')
    root = Tk()
    root.config(border=(20))
    frame = Frame(root, opcao_radio_button=StringVar(), botao_confirma=StringVar())
    root.mainloop()
    

if __name__ == '__main__': main()

