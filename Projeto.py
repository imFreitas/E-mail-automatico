# Importar as Bibliotecas

import pyautogui # Para controlar o mouse, teclado e tela
import pyperclip # Para uso de caracteres especiais 
import time # Para ter um tempo ao rodar o código, evitando bugs

# Passo 01 - Entras no sistema da empresa (no link do drive)

pyautogui.PAUSE = 1
pyautogui.press("win")
pyautogui.write("google chrome")
pyautogui.press("enter")
pyperclip.copy("https://www.linkedin.com/in/imphfreitas/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# Passo 02 - Navegar até o local do relatorio (entrar na pasta Exportar)

time.sleep(3)
pyautogui.doubleClick(x=440, y=387)

# Passo 03 - Exportar o relatório (Fazer o download)

time.sleep(2)
pyautogui.click(x=469, y=487)
time.sleep(2)
pyautogui.click(x=1713, y=195)
time.sleep(2)
pyautogui.click(x=1502, y=556)

# Passo 04 - Calcular os indicadores

tabela = pd.read_excel(r"C:\Users\guiga\Downloads\Vendas - Dez.xlsx")
print(tabela)
faturamento = tabela ["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

# Passo 05 - Enviar o email para a diretoria

pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(3)
pyautogui.click(x=116, y=196)
time.sleep(3)

pyperclip.copy("opedrofreitass@gmail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
pyperclip.copy("Você recebeu um email de Pedro Freitas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")

texto = f"""Olá, Pedro Freitas!
Como pedido, este foi o faturamento da empresa: R${faturamento:,.2f}
E esta foi a quantidade: {quantidade:,}
Este email foi digitado automaticamente!"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("ctrl","enter")
