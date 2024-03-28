from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pyautogui as gui
import pandas as pd
import win32api


driver=webdriver.Edge()
driver.get('https://www.google.com/')

#Criando df e buscando dados 
df=pd.read_excel('lista-logins.xlsx')
for index, row in df.iterrows():
    nome = row["NOME"]
    sobrenome = row["SOBRENOME"]
    d_nascimento = row["D NASCIMENTO"]
    m_nascimento = row["M NASCIMENTO"]
    a_nascimento = row["A NASCIMENTO"]
    genero = row["GENERO"]
    email = row["EMAIL"]
    senha = row["SENHA"]

element=driver.find_element(By.CLASS_NAME, "gb_H")
driver.implicitly_wait(5)   
time.sleep(4)
# gui.click(x=1314, y=263) #click no "fazer login"
login=gui.locateCenterOnScreen('login.PNG',confidence=0.8)
gui.click(login)
time.sleep(15)

#Mes do aniversário
def vef_mes():
    # while m_nascimento != meses:
    #         gui.press('Down', interval=2)
    #         time.sleep(1) 
    # gui.press('enter')
    m = 1
    if m_nascimento == "Janeiro":
        gui.press('Down')
        gui.press('enter')
    elif m_nascimento == "Fevereiro":
        while m < 3:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Março":
        while m < 4:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Abril":
        while m < 5:
            gui.press('Down')
            m+=1
        gui.press('enter') 
    elif m_nascimento == "Maio":
        while m < 6:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Junho":
        while m < 7:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Julho":
        while m < 8:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Agosto":
        while m < 9:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Setembro":
        while m < 10:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Outubro":
        while m < 11:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Novembro":
        while m < 12:
            gui.press('Down')
            m+=1
        gui.press('enter')
    elif m_nascimento == "Dezembro":
        while m < 13:
            gui.press('Down')
            m+=1
        gui.press('enter')

#Genero
def gen_usuario():
    if genero == "Mulher":
        gui.press('Down')
        time.sleep(1)
    elif genero == "Homem":
        gui.press('Down')
        time.sleep(1)
        gui.press('Down')
        time.sleep(1)
    else:
        for i in range(3):
            gui.hotkey('Down', interval=1.2)               



for i in range(4):
    gui.hotkey('Tab', interval=1.2)
gui.press('enter')
time.sleep(1)
gui.press('enter')
time.sleep(10)
gui.typewrite(f"{nome}")
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.typewrite(f"{sobrenome}")
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.press('enter')
time.sleep(8)
    # gui.hotkey('Tab')
    # time.sleep(1)
    # gui.hotkey('Tab')

d=gui.locateOnScreen('dia.png', confidence=0.6)
gui.click(d)
gui.typewrite(f"{d_nascimento}")
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.press('enter')
time.sleep(1)
vef_mes()
time.sleep(3)
# time.sleep(2)
gui.hotkey('Tab')
gui.typewrite(f"{a_nascimento}")
time.sleep(2)
gui.hotkey('Tab')
time.sleep(2)
gen_usuario() #genero do usuario
time.sleep(2)
for i in range(2):
    gui.hotkey('Tab', interval= 1)
# time.sleep(5)
gui.press('enter')

# for i in range(2):
#     gui.hotkey('Tab', interval=1)
# gui.press('enter')
# time.sleep(8)
# for i in range(32): # <-- for para a página de escolha do e-mail
#     gui.hotkey('Tab')   
# time.sleep(2)
# gui.typewrite(email)                       
# gui.hotkey('Tab')
# time.sleep(1)
# gui.press('enter')
# avancar=gui.locateOnScreen('anan.png', confidence=0.8)
# gui.click(avancar)
#     # gui.hotkey("Tab")
    # gui.press('enter')
# time.sleep(3)
# gui.press('enter')
time.sleep(5)
gui.typewrite(email)
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.press('enter')
time.sleep(3)
gui.typewrite(senha)
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.typewrite(senha)       
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)   
gui.press('enter')    
time.sleep(8)

    # driver.refresh


# gui.write('oi')
# element.submit()
# driver.execute_script(f'document.getElementsByClassName("gb_H")')