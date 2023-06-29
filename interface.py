import PySimpleGUI as sg
from openpyxl import Workbook

sg.theme('DefaultNoMoreNagging')

layout = [
    [sg.Text('Digite o nome:')],
    [sg.Input(key='nome', size=(20,1))],

    [sg.Text('Digite a idade:')],
    [sg.Input(key='idade', size=(20,1))],

    [sg.Button('Ok'), sg.Button('Cancelar')]
]

window = sg.Window('Exemplo', layout)

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'Cancelar':
        break

    if event == 'Ok':
        nome_digitado = values['nome']
        idade_digitada = values['idade']

        # Cria um novo arquivo de planilha
        workbook = Workbook()

        # Seleciona a planilha ativa (por padrão, é a primeira planilha criada)
        sheet = workbook.active

        # Define os cabeçalhos das colunas
        sheet['A1'] = 'Nome'
        sheet['B1'] = 'Idade'

        # Insere os dados na planilha
        sheet['A2'] = nome_digitado
        sheet['B2'] = idade_digitada

        # Salva a planilha em um arquivo
        workbook.save('exemplo.xlsx')

        sg.popup('Dados salvos com sucesso!')

window.close()
