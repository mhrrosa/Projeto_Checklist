import json
import random
from collections import OrderedDict
import re
import string
from time import sleep
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree
import pandas
from unidecode import unidecode
import PySimpleGUI as psg



def armazena_perguntas(file_name_xlsx):
    #criando dataframe
    df = pandas.read_excel(file_name_xlsx)
    #criando dict das perguntas
    dict_perguntas = {}

    #passando por cada item da planilha
    for dados in df.iterrows():
        #declarando variaveis e armazenando seus valores
        id = dados[1]['ID']
        descricao = dados[1]['DESCRICAO']
        responsavel = dados[1]['RESPONSAVEL']
        prioridade_peso = dados[1]['PRIORIDADE']

        #armazenando no dict
        dict_perguntas[descricao] = {
            "id": id,
            "responsavel": responsavel,
            "prioridade": prioridade_peso
        }
    return dict_perguntas

def calcula_aderencia(dict_perguntas):

    lista_perguntas =[]
    for key in dict_perguntas.keys():
        lista_perguntas.append(key)

    # define o tema da tela
    psg.theme('reddit')
    contagem = 0
    layout = [
        [psg.Text(size=(70, 1))],
        [psg.Text(f"{lista_perguntas[contagem]}", size=(300, 2), font='Arial 12', justification='center', key='text',
                  text_color='black')],
        [psg.Radio('SIM', 1, enable_events=True, key='R1'), psg.Radio('NÃO', 1, enable_events=True, key='R2'),psg.Radio('NÃO SE APLICA', 1, enable_events=True, key='R3')],
        [psg.Text(size=(70, 1))],
        [psg.Multiline(size=(100, 5),font='Arial 12', key='textbox', disabled=True, visible=False)],
        [psg.Button('Enviar', border_width=2, size=(15, 1))],

    ]
    window = psg.Window('test', layout, size=(1000, 300), element_justification='c')

    while True:
        # le os valores e os eventos
        event, values = window.read()

        if event is None:
            break
        elif values['R1'] == True:
            window['textbox'].update(disabled=True)
            window['textbox'].update(visible=False)
        elif values['R2'] == True  or values['R3'] == True:
            window['textbox'].update(disabled=False)
            window['textbox'].update(visible=True)

        if event == 'Enviar':
            if values['R2'] == True or values['R3'] == True:
                f.write(f"ID: {dict_perguntas[lista_perguntas[contagem]]['id']}\n")
                f.write(f"DESCRICAO: {lista_perguntas[contagem]}\n")
                f.write(f"RESPONSAVEL: {dict_perguntas[lista_perguntas[contagem]]['responsavel']}\n")
                f.write(f"PRIORIDADE: {dict_perguntas[lista_perguntas[contagem]]['prioridade']}\n")
                f.write(f"JUSTIFICATIVA NFC: {values['textbox']}\n")

            window['textbox'].update("")
            window['textbox'].update(disabled=True)
            window['textbox'].update(visible=False)
            window['R1'].update(False)
            window['R2'].update(False)

            contagem += 1
            try:
                window['text'].update(lista_perguntas[contagem])
            except:
                break

def main():
    # selecionando arquivos
    print('Selecione a planilha')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    dict_perguntas = armazena_perguntas(file_name_xlsx)
    calcula_aderencia(dict_perguntas)

if __name__ == '__main__':
    f = open(f"Arquivo_NFC.txt", "w+", encoding='utf8')
    f.write("-- Arquivo de NFC --\n")
    main()
    f.close()
