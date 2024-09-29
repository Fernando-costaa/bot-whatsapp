import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep
from datetime import datetime

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet']
data_atual = datetime.now().date()

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    numero = linha[1].value
    vencimento = linha[2].value
    conta = linha[3].value

    # Verifica se vencimento é do tipo datetime
    if isinstance(conta, datetime):
        # Verifica se vencimento é a mesma data de hoje
        if conta.date() == data_atual:
            # Verifica e converte o número de telefone
            if isinstance(numero, float):
                numero = str(int(numero))  # Converte float para int e depois para string

            # Formata a mensagem
            mensagem = f'Olá (nome da pessoa), hoje dia {conta.strftime("%d/%m/%Y")}, você tem essa conta {vencimento} para pagar.'

            # Abre o link no WhatsApp Web
            link_mensagem_whatsapp = f'(link para a mensagem direta para o whatssap)'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)

            # Localiza o ícone de envio no WhatsApp Web e fecha a conversa
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(5)