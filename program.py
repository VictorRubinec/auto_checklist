import json
import pandas as pd
import requests
import os
from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

url_data = "https://hpe-my.sharepoint.com/personal/victor_rubinec_hpe_com/_layouts/15/download.aspx?UniqueId=27d1c8d4-8dc1-46c0-81db-95a9a9842988&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvaHBlLW15LnNoYXJlcG9pbnQuY29tQDEwNWIyMDYxLWI2NjktNGIzMS05MmFjLTI0ZDMwNGQxOTVkYyIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMCIsIm5iZiI6IjE3MDg5NzMwOTUiLCJleHAiOiIxNzA4OTc2Njk1IiwiZW5kcG9pbnR1cmwiOiI5Szk2MDk5OE1QL3NLQ0dxNkNzQVJTcDR3MytmODdkaXpnRFJqZitQNFQ0PSIsImVuZHBvaW50dXJsTGVuZ3RoIjoiMTQ5IiwiaXNsb29wYmFjayI6IlRydWUiLCJjaWQiOiJmekQzQ3BXT0N3c0dOc2d2c0crdFpBPT0iLCJ2ZXIiOiJoYXNoZWRwcm9vZnRva2VuIiwic2l0ZWlkIjoiWldNeU5EWm1NbUV0TmpjMVpTMDBOalF6TFRobFl6WXRNREZqTURjMVptSmtNV1V6IiwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJnaXZlbl9uYW1lIjoiVmljdG9yIiwiZmFtaWx5X25hbWUiOiJSdWJpbmVjIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJ0aWQiOiIxMDViMjA2MS1iNjY5LTRiMzEtOTJhYy0yNGQzMDRkMTk1ZGMiLCJ1cG4iOiJ2aWN0b3IucnViaW5lY0BocGUuY29tIiwicHVpZCI6IjEwMDMyMDAyN0Q1Q0VFNTIiLCJjYWNoZWtleSI6IjBoLmZ8bWVtYmVyc2hpcHwxMDAzMjAwMjdkNWNlZTUyQGxpdmUuY29tIiwic2NwIjoiZ3JvdXAucmVhZCBteWZpbGVzLnJlYWQgYWxscHJvZmlsZXMucmVhZCIsInR0IjoiMiIsImlwYWRkciI6IjIwLjE5MC4xNzMuMjQifQ.SHk5DmgASB6J6ro7Tj1rAz9L3WK1PXu18GzuMZR55NU&ApiVersion=2.0"
url_template = "https://hpe-my.sharepoint.com/personal/victor_rubinec_hpe_com/_layouts/15/download.aspx?UniqueId=2c40380b-0eb8-4510-80f2-35f446937c4e&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvaHBlLW15LnNoYXJlcG9pbnQuY29tQDEwNWIyMDYxLWI2NjktNGIzMS05MmFjLTI0ZDMwNGQxOTVkYyIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMCIsIm5iZiI6IjE3MDg5NzMwOTUiLCJleHAiOiIxNzA4OTc2Njk1IiwiZW5kcG9pbnR1cmwiOiJiWis0bGt3SW95UWVpVGdxV2lGQ2tvNU1iWFVYbUxQcmJ1Qmg4REZmVzhzPSIsImVuZHBvaW50dXJsTGVuZ3RoIjoiMTQ5IiwiaXNsb29wYmFjayI6IlRydWUiLCJjaWQiOiJmekQzQ3BXT0N3c0dOc2d2c0crdFpBPT0iLCJ2ZXIiOiJoYXNoZWRwcm9vZnRva2VuIiwic2l0ZWlkIjoiWldNeU5EWm1NbUV0TmpjMVpTMDBOalF6TFRobFl6WXRNREZqTURjMVptSmtNV1V6IiwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJnaXZlbl9uYW1lIjoiVmljdG9yIiwiZmFtaWx5X25hbWUiOiJSdWJpbmVjIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJ0aWQiOiIxMDViMjA2MS1iNjY5LTRiMzEtOTJhYy0yNGQzMDRkMTk1ZGMiLCJ1cG4iOiJ2aWN0b3IucnViaW5lY0BocGUuY29tIiwicHVpZCI6IjEwMDMyMDAyN0Q1Q0VFNTIiLCJjYWNoZWtleSI6IjBoLmZ8bWVtYmVyc2hpcHwxMDAzMjAwMjdkNWNlZTUyQGxpdmUuY29tIiwic2NwIjoiZ3JvdXAucmVhZCBteWZpbGVzLnJlYWQgYWxscHJvZmlsZXMucmVhZCIsInR0IjoiMiIsImlwYWRkciI6IjIwLjE5MC4xNzMuMjQifQ.EER7zo96vZa7x3QZnfdkOlmaSJDDnzoFB48nnM6j3eY&ApiVersion=2.0"

def download_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as f:
            f.write(response.content)
        print("Download completo.")
    else:
        print("Falha ao baixar o arquivo.")

if not os.path.exists('template.xlsx'):
    download_file(url_template, "template.xlsx")
    
if not os.path.exists('data.json'):
    download_file(url_data, "data.json")

app_config = json.load(open('app_config.json'))
data = json.load(open('data.json'))
df = pd.DataFrame(data['cliente'])

janela = Tk()
janela.title("Auto Faturamento")

def buscar_por_cnpj(df, cnpj):
    cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
    for cliente in df.iterrows():
        if cnpj in cliente[1]["CNPJ"]:
            valores = cliente[1]["CNPJ"].split(";")
            print(f"valores: {valores}")
            for valor in valores:
                if cnpj in valor:
                    cliente_encontrado = cliente[1].copy() 
                    cliente_encontrado['CNPJ'] = valor 
                    return pd.DataFrame([cliente_encontrado])
    return pd.DataFrame()

def criar_xlsx(cliente, diretorio_saida):
    template = "template.xlsx"
    arquivo_saida = f"cliente_{cliente['CNPJ'].values[0]}.xlsx"
    
    wb = load_workbook(template)
    ws = wb.active
    
    ws['C4'] = cliente['CNPJ'].values[0]
    ws['G4'] = cliente['nome_cliente'].values[0].title()
    ws['G5'] = cliente['endereco_fisico'].values[0].title()
    ws['G6'] = cliente['estado'].values[0].title()
    ws['G7'] = cliente['cidade'].values[0].title()
    ws['H8'] = cliente['party_id'].values[0]
    
    wb.save(f"{diretorio_saida}/{arquivo_saida}")
    
    print(f'Arquivo {arquivo_saida} gerado com sucesso!')
    
    if not cliente.empty:
        print(f'Cliente encontrado: {cliente["nome_cliente"].values[0]}')
    else:
        print(f'Cliente com CNPJ {cnpj_input} não encontrado.')

# funções da interface
def selecionar_diretorio():
    diretorio_saida_input.delete(0, END)
    diretorio_saida_input.insert(0, filedialog.askdirectory())
    app_config['diretorio'] = diretorio_saida_input.get()
    json.dump(app_config, open('app_config.json', 'w'), indent=4)
    
def salvar():
    diretorio_saida = diretorio_saida_input.get()
    cliente = buscar_por_cnpj(df, cnpj_input.get())
    criar_xlsx(cliente, diretorio_saida)
    messagebox.showinfo("Sucesso", "Arquivo gerado com sucesso!")

# elementos da interface
titulo = Label(janela, text="Busca de Cliente por CNPJ")

cnpj_label = Label(janela, text="CNPJ")
cnpj_input = Entry(janela)

diretorio_saida_label = Label(janela, text="Diretório de saída")
diretorio_saida_var = StringVar()  # Variável de controle para o diretório de saída
diretorio_saida_var.set(app_config['diretorio'])  # Definir o valor inicial com o valor do app_config
diretorio_saida_input = Entry(janela, textvariable=diretorio_saida_var)

selecionar_diretorio_button = Button(janela, text="Selecionar", command=selecionar_diretorio)
salvar_button = Button(janela, text="Salvar", command=salvar)
        
# diagramação dos elementos na interface
titulo.grid(row=0, column=0, columnspan=2)
cnpj_label.grid(row=1, column=0)
cnpj_input.grid(row=1, column=1)
diretorio_saida_label.grid(row=2, column=0)
diretorio_saida_input.grid(row=2, column=1)
selecionar_diretorio_button.grid(row=2, column=2)
salvar_button.grid(row=3, column=0, columnspan=2)

# exibição da interface
janela.mainloop()