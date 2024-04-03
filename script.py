import pandas as pd
from openpyxl import load_workbook
from time import sleep

workbook = pd.read_excel('models\\clients.xlsx', sheet_name='Planilha1')

workbook_model = load_workbook('models\\note-model.xlsx')
sheet_model = workbook_model.active

workbook['Data'] = pd.to_datetime(workbook['Data']).dt.date

for i, row in workbook.iterrows():
    nome = row['Nome']
    valor = f"R${row['Valor']:.2f}"
    tipo_de_transacao = row['Tipo']
    data = row['Data']
    identificacao = row['ID']
    cpf = row['CPF']
    telefone = row['Telefone']
    email = row['Email']

    print(nome, valor, tipo_de_transacao, data, identificacao, cpf, telefone, email)

    sheet_model['O2'] = nome #type: ignore
    sheet_model['O4'] = cpf #type: ignore
    sheet_model['O6'] = email #type: ignore
    sheet_model['O8'] = telefone #type: ignore
    sheet_model['O10'] = tipo_de_transacao #type: ignore
    sheet_model['O12'] = valor #type: ignore
    sheet_model['O14'] = data #type: ignore

    workbook_model.save(f'debit-notes\\debit_note_{identificacao}.xlsx')
    print(f'debit note {identificacao} succeed.')
    
sleep(5)