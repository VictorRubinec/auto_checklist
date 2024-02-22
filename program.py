import json
import pandas as pd
from openpyxl import load_workbook

def buscar_por_cnpj(df, cnpj):
    # caso o CNPJ seja uma string, remover a formatação
    # cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
    
    return df[df['CNPJ'] == cnpj]

data = json.load(open('data.json'))

df = pd.DataFrame(data['cliente'])

cnpj_input = input('Digite o CNPJ do cliente: ')

cliente = buscar_por_cnpj(df, cnpj_input)

if not cliente.empty:
    print(cliente)
    
    template = "template.xlsx"
    arquivo_saida = f"cliente_{cnpj_input}.xlsx"
    
    wb = load_workbook(template)
    ws = wb.active
    
    ws['C4'] = cliente['CNPJ'].values[0]
    ws['G4'] = cliente['nome_cliente'].values[0].title()
    ws['G5'] = cliente['endereco_fisico'].values[0].title()
    ws['G6'] = cliente['estado'].values[0].title()
    ws['G7'] = cliente['cidade'].values[0].title()
    ws['H8'] = cliente['party_id'].values[0]
    
    wb.save(arquivo_saida)
    
    print(f'Arquivo {arquivo_saida} gerado com sucesso!')
else:
    print(f'Cliente com CNPJ {cnpj_input} não encontrado.')