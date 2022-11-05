import json
import random
from collections import OrderedDict
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree
import pandas
from unidecode import unidecode



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
        dict_perguntas[id] = {
            "descricao": descricao,
            "responsavel": responsavel,
            "prioridade": prioridade_peso
        }
    return dict_perguntas

def calcula_aderencia(dict_perguntas):
    print("ATENÇÃO: RESPONDA SOMENTE COM A LETRA S OU N")
    for id in dict_perguntas:
        # declarando variaveis e armazenando seu valores
        resposta = ""
        pergunta = dict_perguntas[id]['descricao']
        responsavel = dict_perguntas[id]['responsavel']
        prioridade = dict_perguntas[id]['prioridade']

        while resposta.upper() != "N" and resposta.upper() != "S":
                resposta = input(f"{pergunta} [S/N]").upper()

        if resposta.upper() == "N":
            justificativa = input(f"insira uma justificativa para a NFC: ")
            f.write(f"ID: {id}\n")
            f.write(f"DESCRICAO: {pergunta}\n")
            f.write(f"RESPONSAVEL: {responsavel}\n")
            f.write(f"PRIORIDADE: {prioridade}\n")
            f.write(f"JUSTIFICATIVA NFC: {justificativa}\n")


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