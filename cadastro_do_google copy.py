from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pyautogui as gui
import pandas as pd

mes_disc={
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12
}
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
m = mes_disc.get(m_nascimento)
def vef_mes():
    for _ in range(m):
        gui.press('Down')
    gui.press('enter')
    

#Genero
def gen_usuario():
    if genero == "Mulher":
        time.sleep(0.5)
        gui.press('Down')
        time.sleep(1)
    elif genero == "Homem":
        gui.press('Down')
        time.sleep(1)
        gui.press('Down')
        time.sleep(1)
    elif genero == "Prefiro não dizer":
        for i in range(3):
            gui.hotkey('Down', interval=1.2)  
    else:
        # print('ERRO')
        # driver.quit()
        # driver.stop_casting()
            gui.hotkey('down')
            time.sleep(0.5)               
            gui.hotkey('down')
            time.sleep(0.5)               
            gui.hotkey('down')
            time.sleep(0.5)               
            gui.hotkey('down')
            gui.press('enter')
            time.sleep(0.5) 
            gui.hotkey('Tab')
            time.sleep(0.5) 
            gui.typewrite(f'{genero}')
            gui.hotkey('Tab')
            time.sleep(0.5)
            for _ in range(3):               
                gui.hotkey('Down', interval=1.2)             
            time.sleep(1)
            gui.hotkey('Tab')
            time.sleep(1)
            gui.hotkey('Tab')
            time.sleep(0.5)
            gui.press('enter')
            

for i in range(4):
    gui.hotkey('Tab', interval=1.2)
gui.press('enter')
time.sleep(1)
gui.press('enter')
time.sleep(10)
gui.typewrite(f'{nome}')
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.typewrite(f'{sobrenome}')
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
gui.typewrite(f'{d_nascimento}')
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.press('enter')
time.sleep(1)
vef_mes()
time.sleep(3)
# time.sleep(2)
gui.hotkey('Tab')
gui.typewrite(f'{a_nascimento}')
time.sleep(2)
gui.hotkey('Tab')
time.sleep(2)
gui.press('enter')  
gen_usuario() #genero do usuario
time.sleep(1)
# for i in range(2):
#     gui.hotkey('Tab', interval= 1)
# # time.sleep(5)
# gui.press('enter')


# for i in range(2):
#     gui.hotkey('Tab',interval=1)
# gui.press('enter')
time.sleep(5)
# for i in range(31): # <-- for para a página de escolha do e-mail
#     gui.hotkey('Tab')                           
# time.sleep(2)
# gui.hotkey('Down')
# time.sleep(1)
# gui.hotkey('Down')
# time.sleep(0.5)
# # gui.press('enter')
# time.sleep(1)
gui.typewrite(f'{email}')                       
gui.hotkey('Tab')   
time.sleep(1)
# gui.press('enter')
# avancar=gui.locateOnScreen('anan.png', confidence=0.8)
# gui.click(avancar)
#     # gui.hotkey("Tab")
    # gui.press('enter')
# time.sleep(3)
# gui.press('enter')
# time.sleep(5)
# gui.typewrite(email)
# time.sleep(1)
# gui.hotkey('Tab')
# time.sleep(1)
gui.press('enter')
time.sleep(3)
gui.typewrite(f'{senha}')
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)                       
gui.typewrite(f'{senha}')       
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)
gui.hotkey('Tab')
time.sleep(1)   
gui.press('enter')    
time.sleep(5)

    # driver.refresh


# gui.write('oi')
# element.submit()
# driver.execute_script(f'document.getElementsByClassName("gb_H")')