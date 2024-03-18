import pandas as pd
import time
import webbrowser

from selenium.webdriver.common.by import By
from tkinter import *
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# Buscar arquivo Excel
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    entry_arquivo.delete(0, END)
    entry_arquivo.insert(END, arquivo)

# Buscar coluna retorna informações da linha 0
def buscar_dados():
    nome_do_arquivo = entry_arquivo.get()
    coluna_selecionada = int(entry_coluna.get())

    # Caminho de leitura e execução do chrome
    df = pd.read_excel(nome_do_arquivo)
    chrome_options = Options()
    chrome_options.add_argument("--new-tab")
    chrome = webdriver.Chrome(executable_path="chromedriver.exe", options=chrome_options)

    # AutomaçãoNavegador
    for index, row in df.iterrows():
        print("Index:", index, "", row[coluna_selecionada])
        chrome.get(url_do_navegador)
        time.sleep(1)

        elemento_texto_nome = chrome.find_element(By.XPATH, '//*[@id="APjFqb"]')
        elemento_texto_nome.clear()  # Limpa o campo antes de inserir o novo valor
        elemento_texto_nome.send_keys(str(row[coluna_selecionada]))  # Converte para string

# Ambiente grafico
root = Tk()
root.title("Selecionar Coluna do Excel Começando por 0")
root.geometry("400x200")

label_arquivo = Label(root, text="Selecione o arquivo Excel:")
label_arquivo.pack()

entry_arquivo = Entry(root, width=50)
entry_arquivo.pack()

button_selecionar = Button(root, text="Selecionar", command=selecionar_arquivo)
button_selecionar.pack()

label_coluna = Label(root, text="Digite o número da coluna:")
label_coluna.pack()

entry_coluna = Entry(root, width=10)
entry_coluna.pack()

button_buscar = Button(root, text="Buscar Dados", command=buscar_dados)
button_buscar.pack()

url_do_navegador = 'https://www.google.com'

root.mainloop()
