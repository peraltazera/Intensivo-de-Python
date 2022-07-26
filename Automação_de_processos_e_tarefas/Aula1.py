import pyautogui
import pyperclip
import time
import pandas as pd
from IPython.display import display

pyautogui.PAUSE = 1

# Passo 0: Abrir Google Chrome
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.click(x=1045, y=476)
time.sleep(1)

print("Abriu o Google Chrome")
print("")

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy(
    "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

print("Entrou no Google Drive")
print("")

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=388, y=304)
time.sleep(1)
pyautogui.click(x=388, y=304, clicks=2)
time.sleep(4)
print("Entrou na Pasta Exportar")
print("")

# Passo 3: Fazer o download do relatório
pyautogui.click(x=388, y=304)
pyautogui.click(x=1709, y=189)
pyautogui.click(x=1467, y=558)
time.sleep(10)

print("Iniciou Donwload do Arquivo")
print("")

table = pd.read_excel(r"C:\Users\Victo\Downloads\Vendas - Dez.xlsx")

print("Arquivo Baixado")
print("")
display(table)
print("")

invoicing = table["Valor Final"].sum()
amount = table["Quantidade"].sum()

print("Pegou os Dados do Arquivo Baixado")
print("")

# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

print("Entrou no Email")
print("")

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=38, y=201)
time.sleep(2)

pyautogui.write("vikingzadacontato@gmail.com")
pyautogui.press("tab")  # seleciona o email
time.sleep(1)

pyautogui.press("tab")  # pula pro campo de assunto
pyperclip.copy("Relatório de Vendas")
time.sleep(1)

pyautogui.hotkey("ctrl", "v")  # escrever o assunto
pyautogui.press("tab")  # pular pro corpo do email
time.sleep(1)

textEmail = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${invoicing:,.2f}
A quantidade de produtos foi de: {amount:,}

Abs
Victor Python    
"""

pyperclip.copy(textEmail)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")  # apertar Ctrl Enter para enviar email

print("Enviou o Email")
