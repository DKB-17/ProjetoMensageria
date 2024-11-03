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

for linha in pagina_contatos.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f"Ola {nome}!, Estou informando que deu certo o teste"

    try:
        link_mensagem_wt = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
        navegador.get(link_mensagem_wt)

        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)

        time.sleep(10)
        navegador.find_element(By.XPATH,
            '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
        time.sleep(10)

    except:
        print(f"Nao foi possivel enviar mensagem para {nome}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone};')
