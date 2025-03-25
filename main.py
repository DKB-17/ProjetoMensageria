import time
from time import sleep

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote

navegador = webdriver.Chrome()
navegador.get('https://web.whatsapp.com')

while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(1)

contatos_df = openpyxl.load_workbook('contatos.xlsx')
pagina_contatos = contatos_df['Planilha1']
link_img = f'C:\\Users\\HomePC\\Desktop\\ProjetoMensageria\\beae38d0-e834-42c0-82ba-2bc00f7a87da.jpeg'

for linha in pagina_contatos.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f"Ola {nome}!, Estou testando o envio de mensagem automatizado, favor ignorar as mensagens"

    try:
        link_mensagem_wt = f"https://web.whatsapp.com/send?phone={telefone}"
        navegador.get(link_mensagem_wt)

        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)

        while len(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button')) < 1:
            time.sleep(1)

        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button').click()

        while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input')) < 1:
            time.sleep(1)

        navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input').send_keys(link_img)

        while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div')) < 1:
            time.sleep(1)

        navegador.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div').click()

        time.sleep(10000)

    except:
        print(f"Nao foi possivel enviar mensagem para {nome}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone};')
