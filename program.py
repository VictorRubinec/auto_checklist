import json
import pandas as pd
from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

app_config = json.load(open('app_config.json'))
data = json.load(open('data.json'))
df = pd.DataFrame(data['cliente'])

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

janela = Tk()
janela.title("Auto Faturamento")

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