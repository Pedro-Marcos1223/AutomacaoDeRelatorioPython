import pandas
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

pyautogui.press("win")
pyautogui.write("Navegador Opera GX")
pyautogui.press("enter")

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# while not pyautogui.locateOnScreen("drive.png", confidence=0.9) :
# 	time.sleep(0.5)

time.sleep(3)

pyautogui.click(x=444, y=369, clicks=2)
time.sleep(1)

pyautogui.click(x=554, y=498)  # clicar no arquivo
pyautogui.click(x=1668, y=239)  # clicar nos 3 pontinhos
pyautogui.click(x=1412, y=738)  # clicar no fazer dowload
time.sleep(5)  # esperar o download acabar

tabela = pandas.read_excel(r"C:\Users\pedro\Downloads\Vendas - Dez.xlsx")
print(tabela)  # display para jupyter ou print para outro editor

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

pyautogui.hotkey("ctrl", "t")

pyperclip.copy("mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

pyautogui.click(x=165, y=249)

pyautogui.write("pedromarcos1223@hotmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write("Email de relatorio de vendas")
pyautogui.press("tab")

texto = f"""
Eaee Diretor Pedro,
Aqui esta o relatorio de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Att, Marcos.
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

arquivo = r"C:\Users\pedro\Downloads\Vendas - Dez.xlsx"
pyautogui.click(x=1376, y=985)
pyautogui.click(x=1497, y=902)
time.sleep(2)
pyperclip.copy(arquivo)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

pyautogui.hotkey("ctrl", "enter")