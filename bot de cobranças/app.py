import openpyxl
import pyautogui
from urllib.parse import quote
import webbrowser
import time
import os
import csv

# Carregar a planilha
workbook = openpyxl.load_workbook("contatos.xlsx")
pagina_clientes = workbook['Planilha1']

# Caminho do arquivo de erros
arquivo_erros = 'erros.csv'

# Apagar o arquivo de erros anterior caso ele exista
if os.path.isfile(arquivo_erros):
    os.remove(arquivo_erros)

# Abrir o arquivo de erros em modo append e criar um writer de CSV com codificação utf-8
with open(arquivo_erros, 'a', newline='', encoding='utf-8') as arquivo:
    writer = csv.writer(arquivo)

    # Escrever o cabeçalho do erro.csv
    writer.writerow(['Nome', 'Telefone'])

    #Começa a ler a planilha a partir da linha numero 2
    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value

        # Validação de dados
        if nome is None or telefone is None or vencimento is None:
            print(f"Dados insuficientes na linha {linha[0].row}. Pulando para a próxima linha.")
            writer.writerow([nome, telefone])
            continue

        # Verifica se o telefone está no formato correto (pelo menos 10 dígitos)
        telefone_str = str(telefone)
        if len(telefone_str) < 10 or not telefone_str.isdigit():
            print(f"Número de telefone inválido para {nome} ({telefone}). Pulando para a próxima linha.")
            writer.writerow([nome, telefone])
            continue

        pix = 'XXX.XXX.XXX-XX'

        # Mensagem personalizada dos wpp
        mensagem = f'Olá {nome}, sua mensalidade vence dia {vencimento.strftime("%d/%m/%Y")}. Favor fazer o pagamento no pix: {pix} CPF. Obrigado e tenha um bom dia!'
        link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"

        try:
            webbrowser.open(link_mensagem_whatsapp)
            time.sleep(20)

            #Depuração
            print("Tentando localizar a seta do WhatsApp (modo claro)...")

            #Localizar seta do modo claro do wpp
            try:
                seta_claro = pyautogui.locateCenterOnScreen('seta_claro.png', confidence=0.9)
                pyautogui.click(seta_claro[0], seta_claro[1])
            except Exception as e:
                print("Não foi possível localizar a seta no modo claro. Tentando modo escuro...")
                try:
                    seta_escuro = pyautogui.locateCenterOnScreen('seta_escuro.png', confidence=0.9)
                    pyautogui.click(seta_escuro[0], seta_escuro[1])
                except:
                    print('Não foi possível encontrar nenhuma das setas')
                    raise Exception('ERRO')
            time.sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(5)

        except Exception as e:
            print(f"Não foi possível enviar a mensagem para {nome} ({telefone}): {e}")
            writer.writerow([nome, telefone])
