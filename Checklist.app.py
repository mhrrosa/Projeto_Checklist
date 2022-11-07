
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas
import PySimpleGUI as psg
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


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
    quantidade_aprovados = 0
    quantidade_total = 0
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
        [psg.Button('Enviar',image_data=download3, button_color=('white', psg.theme_background_color()), border_width=0, )]

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
                f.write(f'-----------------\n')

            else:
                quantidade_aprovados +=1
            quantidade_total +=1
                

            window['textbox'].update("")
            window['textbox'].update(disabled=True)
            window['textbox'].update(visible=False)
            window['R1'].update(False)
            window['R2'].update(False)

            contagem += 1
            try:
                window['text'].update(lista_perguntas[contagem])
            except:
                f.write(f'\n')
                f.write(f'\n')
                f.write(f'--- Estatísticas ---\n')
                f.write(f'\n')
                f.write(f'- Quantidade aprovados: {quantidade_aprovados}\n')
                f.write(f'- Quantidade total: {quantidade_total}\n')
                f.write(f'- Taxa de aderencia: {quantidade_aprovados/quantidade_total *100}%  - \n')
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
    f.write(f'\n')
    download3 = b'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/GMzLdsAAFTiaVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/Pgo8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA1LjYtYzE0NSA3OS4xNjIzMTksIDIwMTgvMDIvMTUtMjA6Mjk6NDMgICAgICAgICI+CiAgIDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+CiAgICAgIDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiCiAgICAgICAgICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgICAgICAgICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIKICAgICAgICAgICAgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iPgogICAgICAgICA8eG1wOkNyZWF0b3JUb29sPkFkb2JlIFBob3Rvc2hvcCBFbGVtZW50cyAxNy4wIChXaW5kb3dzKTwveG1wOkNyZWF0b3JUb29sPgogICAgICAgICA8eG1wOkNyZWF0ZURhdGU+MjAyMC0xMC0wM1QxMTowOTowOS0wNDowMDwveG1wOkNyZWF0ZURhdGU+CiAgICAgICAgIDx4bXA6TWV0YWRhdGFEYXRlPjIwMjAtMTAtMDNUMTE6MDk6MDktMDQ6MDA8L3htcDpNZXRhZGF0YURhdGU+CiAgICAgICAgIDx4bXA6TW9kaWZ5RGF0ZT4yMDIwLTEwLTAzVDExOjA5OjA5LTA0OjAwPC94bXA6TW9kaWZ5RGF0ZT4KICAgICAgICAgPHhtcE1NOkluc3RhbmNlSUQ+eG1wLmlpZDowNDFhMDAwNy1lMGJlLTk4NGYtODNlMy01ZGQxZTk2NjI5ZWQ8L3htcE1NOkluc3RhbmNlSUQ+CiAgICAgICAgIDx4bXBNTTpEb2N1bWVudElEPmFkb2JlOmRvY2lkOnBob3Rvc2hvcDo1YjgyOTQ2Mi0wNThhLTExZWItOTQ3ZC04N2E5Njc3OWZkYzU8L3htcE1NOkRvY3VtZW50SUQ+CiAgICAgICAgIDx4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ+eG1wLmRpZDoxMWMzOWY4MS0xM2VmLWIzNGMtYWNkMy0xZTVjMTI5OWNmMGM8L3htcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD4KICAgICAgICAgPHhtcE1NOkhpc3Rvcnk+CiAgICAgICAgICAgIDxyZGY6U2VxPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5jcmVhdGVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6MTFjMzlmODEtMTNlZi1iMzRjLWFjZDMtMWU1YzEyOTljZjBjPC9zdEV2dDppbnN0YW5jZUlEPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6d2hlbj4yMDIwLTEwLTAzVDExOjA5OjA5LTA0OjAwPC9zdEV2dDp3aGVuPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBQaG90b3Nob3AgRWxlbWVudHMgMTcuMCAoV2luZG93cyk8L3N0RXZ0OnNvZnR3YXJlQWdlbnQ+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5zYXZlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOjA0MWEwMDA3LWUwYmUtOTg0Zi04M2UzLTVkZDFlOTY2MjllZDwvc3RFdnQ6aW5zdGFuY2VJRD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OndoZW4+MjAyMC0xMC0wM1QxMTowOTowOS0wNDowMDwvc3RFdnQ6d2hlbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgUGhvdG9zaG9wIEVsZW1lbnRzIDE3LjAgKFdpbmRvd3MpPC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICAgICA8c3RFdnQ6Y2hhbmdlZD4vPC9zdEV2dDpjaGFuZ2VkPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICA8L3htcE1NOkhpc3Rvcnk+CiAgICAgICAgIDxwaG90b3Nob3A6RG9jdW1lbnRBbmNlc3RvcnM+CiAgICAgICAgICAgIDxyZGY6QmFnPgogICAgICAgICAgICAgICA8cmRmOmxpPjAwMDE1N0JCNEVDNjZDODMyQ0VBN0Q0OTgxOEYyQkI3PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+MEQyMERDMEVCMTBDQkE5Njg5N0M2NzNCRjkwNDI5ODQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT4xRDk5RjMzMUY1RkMyOUU0ODU5MkI1OERENENCRkUzMzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPjI2QzMwRDNBRTREQjZERTFFN0Y2M0JCQUE4NjBGNEI0PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+MkFFRkE4NTk0ODJBRTMxMEYwOEYxNEVCQkU3MUEyNTU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT4zNDBDQUZCNkZCMzIwRDRGREVEMjc0M0ExRjUwNUI2ODwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPjQ4NTVBMzI3NzUwOTIzODkwMzQ5NjIwRkU2NUYzNjkxPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+NjZBREMzOThERjcwMDQ1RDgxMkU4OUMwNDIzRkFGNTA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT42QjI5MkM4MDQyRTY1QTcxMkZGMTk4NTdEMjhGQTZCRTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPjczNzQ2N0JGQzU2QkFDNTk2Q0M4QkNEOUUzNjk2QUU1PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+NzNDMEVENEM3ODE0RTg4RjlBMjQ3NzRFRjdGMTBBODk8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT44MUU0QzY3QjQ5QkFCMzlDNkU5QzExRjQxNUNEMTgyRTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPkEzQ0M2RjFEQTFDMTFBMDZDOUExOTEyQURDRDBFRjQ3PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+QTQ3QkZCNkY5NkMxRjBGMjhDMTI5RENCQkZBODRGNkQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT5BNTlDQUU4ODNCMzU0RjgyMEQ3OTFEODJCREVGRjE2MDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPkFGREUyQkMyMzA1QkUyRTc2Q0RCNTdBNDAwNzM3MEQyPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+Q0Q4RTcxRkQ1REQ3RkM5MjcyNkZFREQ5NDRBREEyMTE8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT5GM0Y5OURGQjBFMzE5QzUwQzRGNEQ2NUZCM0U1QjU5MzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPkY4REE4QjZDRUM5MDI5OTgzMzUxQkEzQzUyQTVCNzREPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+YWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOjMyMjNhYzg2LWVkZTYtMTFkOS1hZGRiLWNiMWQ2NjAxOWMxNzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPmFkb2JlOmRvY2lkOnBob3Rvc2hvcDo2YjM0YmVlZi1mNDg2LTExZDktOWFiYy1mNmY3MDc0YjFkN2E8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT5hZG9iZTpkb2NpZDpwaG90b3Nob3A6NmNjNzJkZDQtYzEyMi0xMWRhLTllYTAtYjQxMDIxN2JjNjA0PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+YWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOjc4ZTg3ZmZlLWY5NWEtMTFkYi1hZmE5LWY3OGRhN2FiODZmZjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPmFkb2JlOmRvY2lkOnBob3Rvc2hvcDo5NWFkMzRlOC01MTQwLTExZGEtOGFmNi1iNjQ4NmE1YjIwYjI8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT5hZG9iZTpkb2NpZDpwaG90b3Nob3A6YmViNmY0N2UtMjIzNS0xMWRjLWE0NGUtZjZiOGI4MzA4MWQ1PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+YWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOmU3ZmI1OGNhLTk2NDItMTFkZC04MWEyLWRkMTQxMDNmNDUxNTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6MDFGMUEwRjYxMUVGMTFEQjg4MUVBNkJFMDVFMjc2RDE8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOjA2MDFFMzM5QkFGMkREMTE5QTlEOTEwOTg5NjI4QzVGPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDoxODkyOTc1OTlFNUREQzExOEQxOUJGODREMUU2QjZEQTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6MTk1MDE1MDhFMkVGMTFERDhCRUNDQjZCNDU1MkJFQTc8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOjJBNTJBMDEwMENGNERDMTFBQTAzRUZBRjc4RDI5ODlDPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDozMEMyMzlCODJFM0JERDExQjY1MUMwNzNDODY5RDI3MzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6MzUwMEIwQ0M2QzUwMTFEQ0E1OEE4MUNDRUFFQjk2N0Q8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOjRBMzhCMjREN0NCNkREMTE5ODk5OUZCM0IyMkVFNUNFPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDo1MDg1RkUxQzVDRjFEQzExODQ5OTlCMkQ0NzNCNDBDNTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6NTY5OTMxQUNFRTgyREMxMTk4NkVFQTgzNjFGMTQ3MTE8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOjU3RDRENjNDMDUzN0UwMTFBNjM1RUJFMzgyMzBCODQxPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDo1RjgzOTA1RDExRUYxMURCODgxRUE2QkUwNUUyNzZEMTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6NjI1OTE1RTFEMEYyREQxMUIyQUVGODg5ODVDQTU4Njg8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOjcwNkY4OEYyMzAzQkREMTFCNjUxQzA3M0M4NjlEMjczPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDo5MEFERjdENTM5OEExMURGODYxMTk0MzZBMTdDQjIyNzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6QTFFMjhDNEM1NjI5REMxMTlDNzZCNzZDOEFBMDk4NzQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOkE5Rjg5MTI5M0IxQzExREFCNTVCQzI1REIzOTc4NjA2PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDpCNTdGQjJFRUM4MEJEQzExOURCQThEMjJDNkE3OUM0NzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6QkQ1NkJGM0FGQzM4REIxMUEwNzhFRjBBNDMwOTAyRDU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOkJFREZCQTY5MEM2OERCMTE5NEE0RkFCMjk1MERCMjQ1PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDpDQUQxM0U4ODMwM0JERDExQjY1MUMwNzNDODY5RDI3MzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnV1aWQ6RDlGRkY3NkJDMThBRTAxMTk3RTJFNkIyMDZFQUIzOTU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT51dWlkOkU0NUNENDFBNzdEN0REMTE4NkRFREQ1OEI3N0JFMDkxPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+dXVpZDpGNkFBNkUxMThCNURFMTExOTE0MUJBNTI3OTk1MzhGMDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6MDE4MDExNzQwNzIwNjgxMTg3MUZFNEFENzFFRTI1QzA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjAxODAxMTc0MDcyMDY4MTE4NzFGRUVCNTZEMUU3NDNCPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDowMTgwMTE3NDA3MjA2ODExOEY2MkVCMkFERTQ2RDI3MzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6MDE4MDExNzQwNzIwNjgxMTkxMDlGQTI5N0E3QTU5MDQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjAxODAxMTc0MDcyMDY4MTE5MkIwQkFBOTA0REUwRjhEPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDowMTgwMTE3NDA3MjA2ODExOTdBNURBRjI1ODNBMEE0QjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6MDE4MDExNzQwNzIwNjgxMUFCMDhFOEU4RUUzRjAyODk8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjAyODAxMTc0MDcyMDY4MTE4OEM2QTI5ODQ0M0YxMEFDPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDowMjgwMTE3NDA3MjA2ODExOTEwOUM2NUE3MDQwMTM0MDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6MEEwOEJDOTI2MUEyREYxMUJCQjVDMUZFMzRDNjY0OEM8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjFDOEU5N0FBOEI0QkRGMTE4OTQxQjEwMTZEQkYxRDE0PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDoxRjgzQ0M1NjM2RTFFMDExQjc1MzgyQTY2NTdEMjA4RDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6MzE5RTE0NkRCNDM3RTIxMThENjBEQkI0MTU3NjYyQkU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjM2RThEMTc2QkM3QURGMTFBM0M4QzZFMjExRUE4QkJGPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDo0RDVFODcxRTE1MTFFMjExOTU1M0RENDcyOUEyMzZDRjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6NTE0MDRGNDMzOTQ0RTIxMUEyMzlDMDRBRkQ0NDQ4MUU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjUxNTQ2RkREMTMyMDY4MTE5MkIwQkFBOTA0REUwRjhEPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDo1NDU0NkZERDEzMjA2ODExOTJCMEJBQTkwNERFMEY4RDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6NjM4ODgwRTgwOTIwNjgxMTkyQjBCQUE5MDRERTBGOEQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjY0ODg4MEU4MDkyMDY4MTE5MkIwQkFBOTA0REUwRjhEPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDo2NERGNTM3NjM1RTFFMDExQjc1MzgyQTY2NTdEMjA4RDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6NjY4ODgwRTgwOTIwNjgxMTkyQjBCQUE5MDRERTBGOEQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjY3ODg4MEU4MDkyMDY4MTE5MkIwQkFBOTA0REUwRjhEPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDo2RUQ1RTlDRDZGQzRERjExQjc1QzlCRTFENUIxNkVBMzwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6NzNGREFEODU3RUNCRTExMUI1MUVCRkE3RDIwMUNDNEI8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOjk2MkFGRDIxNjgzOUUyMTE4NzA1ODUzNjA1RUY2MEJBPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpBMzQ3Q0E3MUNFRENFMTExOUJFN0Y0NzJGQzQyQzU4NTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6QTY1Q0FDMUZDOTM3RTIxMThENjBEQkI0MTU3NjYyQkU8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOkE5Q0NBNkFGRkI0M0UxMTE5RURCRjYzRTA1ODA0MzA0PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpCMjQwMUYyQjk2MzhFMjExQTY2MkI0N0U5QjhENzE0RTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6QkIzNTY1RTgxNTIwNjgxMTkyQjBCQUE5MDRERTBGOEQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOkMyMzU2NUU4MTUyMDY4MTE5MkIwQkFBOTA0REUwRjhEPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpDMzczQ0VDMDM2MjA2ODExOEY2MkQwRjcwMTBBQzAyRjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6Q0I1NzY4QzhGQjNBRTIxMUFBNDNGMUNGRTNEMkUyNTM8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOkNGMEY3NzJFOUY5MkUwMTFCMDBBQjc4N0Y4ODIyQjQ2PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpEMDczREU4OURFM0VFMjExQkM2QUZERjZBNEZEMTc1RjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6RDQwNTRFQzExOTIwNjgxMTkyQjBCQUE5MDRERTBGOEQ8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOkRBQTQ4QUM0NzYzRkUyMTE4M0EyQkZEMTQ2NjBDQTJGPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpFMEE0OEFDNDc2M0ZFMjExODNBMkJGRDE0NjYwQ0EyRjwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6Rjc1RDkxOEExMTIwNjgxMTkyQjA4QkVFMjlDNzVERDI8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaT54bXAuZGlkOkY3N0YxMTc0MDcyMDY4MTE4MDgzRUI4M0M2MkJEN0MxPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+eG1wLmRpZDpGNzdGMTE3NDA3MjA2ODExOTEwOUY4RkUyNzcxOEQ1QTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPnhtcC5kaWQ6RkJDNEQyMDQwQTIwNjgxMTkxMDlDQzY0MkM0NEVDMEM8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6QmFnPgogICAgICAgICA8L3Bob3Rvc2hvcDpEb2N1bWVudEFuY2VzdG9ycz4KICAgICAgICAgPHBob3Rvc2hvcDpDb2xvck1vZGU+MzwvcGhvdG9zaG9wOkNvbG9yTW9kZT4KICAgICAgICAgPHBob3Rvc2hvcDpJQ0NQcm9maWxlPnNSR0IgSUVDNjE5NjYtMi4xPC9waG90b3Nob3A6SUNDUHJvZmlsZT4KICAgICAgICAgPGRjOmZvcm1hdD5pbWFnZS9wbmc8L2RjOmZvcm1hdD4KICAgICAgICAgPHRpZmY6T3JpZW50YXRpb24+MTwvdGlmZjpPcmllbnRhdGlvbj4KICAgICAgICAgPHRpZmY6WFJlc29sdXRpb24+NzIwMDAwLzEwMDAwPC90aWZmOlhSZXNvbHV0aW9uPgogICAgICAgICA8dGlmZjpZUmVzb2x1dGlvbj43MjAwMDAvMTAwMDA8L3RpZmY6WVJlc29sdXRpb24+CiAgICAgICAgIDx0aWZmOlJlc29sdXRpb25Vbml0PjI8L3RpZmY6UmVzb2x1dGlvblVuaXQ+CiAgICAgICAgIDxleGlmOkNvbG9yU3BhY2U+MTwvZXhpZjpDb2xvclNwYWNlPgogICAgICAgICA8ZXhpZjpQaXhlbFhEaW1lbnNpb24+MTA5PC9leGlmOlBpeGVsWERpbWVuc2lvbj4KICAgICAgICAgPGV4aWY6UGl4ZWxZRGltZW5zaW9uPjM4PC9leGlmOlBpeGVsWURpbWVuc2lvbj4KICAgICAgPC9yZGY6RGVzY3JpcHRpb24+CiAgIDwvcmRmOlJERj4KPC94OnhtcG1ldGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgCjw/eHBhY2tldCBlbmQ9InciPz4BWge9AAAAIGNIUk0AAHolAACAgwAA+f8AAIDpAAB1MAAA6mAAADqYAAAXb5JfxUYAAAHGSURBVHja7Nu9ihNRHIbx5/xnJpOZaAKbXWIjwUULkTU2rqgoYqM2FnaWfqAXYGWpjbjVegFiue028aPbzgURFuxUtNMNYkSLOJsx59joPZzA+1zC++PMnOa4EAKA3Xq0ceLd59HaZDo746FARVHiqIos2V451Lv39P71HcC7EII9ePbq+Obr91sha3aSrIEz01qRFLxnVk+hrn5eXT188eHtKzsp0Hrzcfexa+7rNIoSS1JwTmtFoxbwsz/UVdp5+2m0BlwzoPfjtz+bFSWWZgKLLeewNCNrlown/jTQM6BbzVzLJakGitkuSZliJbBgQO6SBKcTFvmBczhLAHIDdOuYr3RNnEs1TSA0JTQlNKEpoSmhCU0JTQlNaEpoSmhKaEJTQlNCE5oSmhKa0JTQlNCEpuYIzbfLfKop4m9/0agBb8DeoN/98u9xoYq0EAIr/e5XYM+A73cuD4bt3GaaJt7aufm7lwZDYGzA6OjBxedPbp5/sbrcrZqpfnMxlafGyeXFav3GuZfH+ktDYORCCAYsAaeAC8ARoAXo7VMEX0VgAnwAtoBt4Jv7/1AeaAMHgAUgE1o0aDUwBnaBX4D/CwAA//8DAOgHcZ21310BAAAAAElFTkSuQmCC'

    main()
    f.close()
