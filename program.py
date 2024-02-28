import json
import pandas as pd
import requests
import os
from openpyxl import load_workbook
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# ========================================================================================================
# Configuração
# ========================================================================================================

app = {
    "App_Name" : "Auto Faturamento",
    "App_Version" : "1.0",
    "App_Logo" : "logo.png",
    "App_Author" : "Victor Zanin Rubinec",
    
    "Conf_Geometry" : "600x225+100+100",
    
    "Header_color" : "#1aae9f",
    "Footer_color" : "#4b5c6b",
    
    "Buttom_head_color" : "#e3e8ed",
    "Buttom_salvar_color" : "#d5e7f7",
    "Buttom_diretorio_color" : "#d1efec",
    
    "Buttom_head_color_active" : "#b7c7cf",
    "Buttom_salvar_color_active" : "#74b0e6",
    "Buttom_diretorio_color_active" : "#3bbaad",
    
    "titulo_color" : "#293845",
    
    "titulo" : "Auto Faturamento",
    "cpnj_label" : "Digite o CNPJ desejado para o documento:",
    
    "Buttom_salvar" : "Salvar",
    "Buttom_diretorio" : "Diretório destino",
    "Buttom_ajuda" : "Ajuda",
    "Buttom_sobre" : "Sobre",
}

# Urls dos arquivos que estão no onedrive
url_data = "https://graph.microsoft.com/v1.0/me/drive/root/children/Auto_Financiamento/children"
token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6Im90Nkd0QjBxN0N3NEY4ZzBlaTlDejR3aEZhR0lLem0zaVQzSmtSS0xPelkiLCJhbGciOiJSUzI1NiIsIng1dCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSIsImtpZCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8xMDViMjA2MS1iNjY5LTRiMzEtOTJhYy0yNGQzMDRkMTk1ZGMvIiwiaWF0IjoxNzA5MDU2Mzk4LCJuYmYiOjE3MDkwNTYzOTgsImV4cCI6MTcwOTE0MzA5OCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFUUUF5LzhXQUFBQTJsd1lCY2dWQlk4NXcwcFZvb3VZTFRSanUzS0Jtc3orYnpkS3hNdFo5UE50NlQ2MHNMU28rczlaaUxkazRTUTQiLCJhbXIiOlsicnNhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IlJ1YmluZWMiLCJnaXZlbl9uYW1lIjoiVmljdG9yIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTc5LjI0NS43MC4xNDIiLCJuYW1lIjoiUnViaW5lYywgVmljdG9yIiwib2lkIjoiNjcyMDkyZGMtOTFhNy00ZGViLWE3YjQtZjI0ZjdmMWQxMWQwIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTgzOTUyMjExNS0xMzgzMzg0ODk4LTUxNTk2Nzg5OS02NzE4Mjc5IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAyN0Q1Q0VFNTIiLCJyaCI6IjAuQVMwQVlTQmJFR20yTVV1U3JDVFRCTkdWM0FNQUFBQUFBQUFBd0FBQUFBQUFBQUF0QUdRLiIsInNjcCI6IkFQSUNvbm5lY3RvcnMuUmVhZC5BbGwgQVBJQ29ubmVjdG9ycy5SZWFkV3JpdGUuQWxsIENoYXQuUmVhZFdyaXRlIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkLkFsbCBEaXJlY3RvcnkuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkIEdyb3VwLlJlYWQuQWxsIG9wZW5pZCBPcmdDb250YWN0LlJlYWQuQWxsIFBlb3BsZS5SZWFkIHByb2ZpbGUgUmVwb3J0cy5SZWFkLkFsbCBVc2VyLlJlYWQgZW1haWwiLCJzdWIiOiIxd1dHRFJsR3VBNUVHZUljY1BQdUtIZElZSG1BQzFpMDFNcTZwVXpiVVJnIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiMTA1YjIwNjEtYjY2OS00YjMxLTkyYWMtMjRkMzA0ZDE5NWRjIiwidW5pcXVlX25hbWUiOiJ2aWN0b3IucnViaW5lY0BocGUuY29tIiwidXBuIjoidmljdG9yLnJ1YmluZWNAaHBlLmNvbSIsInV0aSI6Imx1QTZydDZGXzBXaEpaRGk1eGd4QUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfY2MiOlsiQ1AxIl0sInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6Imt6M2g3UDJ4elFxaXd5c0xrR3VfdEJzUnVWbHhWOFRQa3EtbW94ZXY5Qk0ifSwieG1zX3RjZHQiOjE0MTY2MTUyMTl9.jQ9q42XS3ZIyywwsBoPxK7Oxlwtbg4SofQ5POAPeRZ5EfIASnqNBB6HbVppC1I_Kuri-a3PA8eYK7Fbbfw8vQurYSvjX5dRPqok9Mj7Avu6MOqtKoKrNs4EBRg0Tv-Z73U7ws1RnD4tqDBWIynyh2lBqAyrvUWJHhdZw2nXBkvRxo25KKfHZ6HRj9pgCWYkezSeDjZdf3NRw0sgpN8chJrwMqmiAk0z3a6Wn1FMpHuTcileBYU_Wqelk9wkIQS45MhJX4DK3cgGugB0kfAI7D2d3ynYyD2VqWI3zgM-ZdsHzK-TBQm3wRiEGvyIhN06imDGHn_tpAbIuK22l7LeVDA"

# Função para baixar arquivos do onedrive
def download_file(url, filename):
    with open(filename, 'wb') as f:
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"})
        arquivos = pd.DataFrame(response.json()['value'])
        for item in arquivos.iterrows():
            if (item[1]['name'] == filename):
                url = item[1]['@microsoft.graph.downloadUrl']
                response = requests.get(url)
                f.write(response.content)
        if response.status_code == 200:
            print(f"Arquivo {filename} baixado com sucesso!")
        else:
           print(f"Erro ao baixar o arquivo {filename}!")

# Verificando se os arquivos já existem, caso não existam, baixar ou criar
if not os.path.exists('template.xlsx'):
    download_file(url_data, "template.xlsx")    
if not os.path.exists('data.json'):
    download_file(url_data, "data.json")
if not os.path.exists('logo.png'):
    download_file(url_data, "logo.png")
if not os.path.exists('app_config.json'):
    app_config = {"diretorio": ""}
    json.dump(app_config, open('app_config.json', 'w'), indent=4)

# Carregando os dados em variáveis
app_config = json.load(open('app_config.json'))
data = json.load(open('data.json'))
df = pd.DataFrame(data['cliente'])

# ========================================================================================================
# Funções
# ========================================================================================================

# Função para buscar cliente por CNPJ
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

# Função para criar o arquivo xlsx
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
        print(f'Cliente com CNPJ {input} não encontrado.')

# Função para selecionar o diretório de saída do arquivo
def selecionar_diretorio():
    diretorio_saida = filedialog.askdirectory()
    diretorio_label.config(text=diretorio_saida)
    app_config['diretorio'] = diretorio_saida
    json.dump(app_config, open('app_config.json', 'w'), indent=4)
    
# Função para salvar o arquivo no diretório selecionado
def salvar():
    if app_config['diretorio'] == "":
        messagebox.showerror("Erro", "Selecione um diretório de saída!")
        return
    diretorio_saida = app_config['diretorio']
    cliente = buscar_por_cnpj(df, input.get())
    criar_xlsx(cliente, diretorio_saida)
    messagebox.showinfo("Sucesso", "Arquivo gerado com sucesso!")

# ========================================================================================================
# Interface
# ========================================================================================================

# Criação da janela
janela = Tk()
janela.title(app["titulo"])
janela.geometry(app["Conf_Geometry"])

# Elementos da interface
head = Frame(janela, width=600, height=75, background=app["Header_color"])
body = Frame(janela, width=600, height=100)
footer = Frame(janela, width=600, height=50, background=app["Footer_color"])

titulo = Label(head, text=app["titulo"], font=("Arial", 27, "bold"), background=app["Header_color"], foreground=app["titulo_color"], padx=10)

cnpj_label = Label(body, text=app["cpnj_label"], font=("Arial", 12))
diretorio_label = Label(footer, text=app_config["diretorio"], font=("Arial", 10), background=app["Footer_color"], foreground="white")

input = Entry(body, font=("Arial", 12), border=1, cursor="xterm", width=30)

btn_grid_head = Frame(head, width=100, height=75, background=app["Header_color"])
btn_grid_body = Frame(body, width=100, height=100)

btn1 = Button(btn_grid_body, text=app["Buttom_salvar"], font=("Arial", 12), background=app["Buttom_salvar_color"],
                padx=15, pady=5, border=1, activebackground=app["Buttom_salvar_color_active"], cursor="hand2", command=salvar)
btn2 = Button(btn_grid_body, text=app["Buttom_diretorio"], font=("Arial", 12), background=app["Buttom_diretorio_color"],
                padx=15, pady=5, border=1, activebackground=app["Buttom_diretorio_color_active"], cursor="hand2", command=selecionar_diretorio)
btn3 = Button(btn_grid_head, text=app["Buttom_ajuda"], font=("Arial", 12), background=app["Buttom_head_color"],
                padx=15, pady=5, border=1, activebackground=app["Buttom_head_color_active"], cursor="hand2")
btn4 = Button(btn_grid_head, text=app["Buttom_sobre"], font=("Arial", 12), background=app["Buttom_head_color"],
                padx=15, pady=5, border=1, activebackground=app["Buttom_head_color_active"], cursor="hand2")

logo_dir = os.getcwd() + "\\" + app["App_Logo"]
logo = PhotoImage(file=logo_dir)
logo_label = Label(footer, image=logo, background=app["Footer_color"])
        
# Diagramação dos elementos na interface
head.pack(side="top")
head.pack_propagate(False)
head.grid_propagate(False)

titulo.pack(side="left")

# ainda sem os documentos dos botões
# btn_grid_head.pack(side="right")
# btn3.grid(row=0, column=1, padx=5)
# btn4.grid(row=0, column=2, padx=5)

body.pack(side="top")
body.grid_propagate(False)
body.pack_propagate(False)

cnpj_label.grid(row=0, column=0, padx=5, pady=5)
input.grid(row=1, column=0, padx=15)

btn_grid_body.grid(row=2, column=1, padx=15)
btn2.grid(row=0, column=0, padx=5)
btn1.grid(row=0, column=1, padx=5)

footer.pack(side="bottom")
footer.grid_propagate(False)
footer.pack_propagate(False)

diretorio_label.grid(row=0, column=0, padx=2, pady=2)
logo_label.pack(side="right")

# Exibição da interface
janela.mainloop()