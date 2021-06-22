#!/usr/bin/env python
# coding: utf-8

# In[16]:


import pyautogui
import time
import pyperclip

pyautogui.PAUSE=1
#passo 1
#abrir uma nova aba
pyautogui.hotkey("ctrl", "t")
#entrar no link do sistema
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
#---------------------------------

#passo 2
time.sleep(2)
pyautogui.click(405, 276, clicks=2)
#------------------------------------

#passo 3
pyautogui.click(416, 365)
pyautogui.click(1159, 183)
time.sleep(1)
pyautogui.click(987, 542)
time.sleep(3)
pyautogui.click(1040, 555)




# In[24]:


import pandas as pd
#passo 4
tabela = pd.read_excel(r"C:\Users\55119\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum() #soma da coluna valor final
quantidade = tabela["Valor Unitário"].sum()#soma da coluna quantidade
display(tabela)


# In[45]:


#passo 5
pyautogui.hotkey("ctrl", "t")
gmail="https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin"
detisnatario="semchao342@gmail.com"
assunto="Relatório de Vendas"
texto_email=f""" 
Prezados, Bom dia

O faturamento de ontem foi : R${faturamento, .2f}.
A quantidade de produtos vendidos foi : {quantidade} Itens.

Abs
Yuri Costa Camilo
"""
pyperclip.copy(gmail)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("enter")
time.sleep(4)
pyautogui.click(101,196)
time.sleep(1)
pyperclip.copy(detisnatario)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyperclip.copy(texto_email)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)
pyautogui.hotkey("ctrl", "enter")













# In[40]:


time.sleep(5)
pyautogui.position()


# In[ ]:




