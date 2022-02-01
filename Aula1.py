import pyautogui as pag
import pyperclip as pyp
import pandas as pd
import time

pag.PAUSE = 2


pag.hotkey("win", "s")
pag.write("Chrome")
pag.press("Enter")

pyp.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pag.hotkey("ctrl", "v")
pag.press("enter")

pag.click(x = 400, y = 300, clicks = 2)
pag.click(x = 400, y = 350, button = "right")
pag.click(x = 544, y = 864)

df = pd.read_excel(r"C:\Users\lukef\Downloads\Vendas - Dez.xlsx")
valor_vendas = df["Valor Final"].sum()
unidades_vendas = df["Quantidade"].sum()

pag.hotkey("ctrl", "t")
pyp.copy("https://mail.google.com/mail/u/0/#inbox")
pag.hotkey("ctrl", "v")
pag.press("enter")

pag.click(x = 87, y = 209)
time.sleep(2)
pyp.copy("lukefreitasberry@gmail.com")
pag.hotkey("ctrl", "v")
pag.press("tab")
pyp.copy("Relatório Mensal - Dezembro")
pag.hotkey("ctrl","v")

pag.press("tab")
pag.write(
   f""" 
Boa noite, Luke!

Seguem os resultados de Dezembro como discutido:

Unidades Vendidas: {unidades_vendas} unidades
Valor Total das Vendas: R${valor_vendas:,.2f}
   """
)

pag.press("tab")
pag.press("enter")
pag.hotkey("alt", "F4")
