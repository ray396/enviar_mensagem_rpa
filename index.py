from selenium import webdriver
import pyautogui as tempoEspera
import pyautogui as teclasTeclado

from selenium.webdriver.common.by import By

from openpyxl import load_workbook

nome_arquivo_contatos = ".\WhatsApp\Contatos.xlsx"

planilhaDadosContato = load_workbook(nome_arquivo_contatos)

sheet_selecionada = planilhaDadosContato['Dados']

navegadorChrome = webdriver.Chrome()

navegadorChrome.get("https://web.whatsapp.com/")

tempoEspera.sleep(25)

while len(navegadorChrome.find_elements(By.ID,'side')) < 1:
    tempoEspera.sleep(3)

tempoEspera.sleep(4)

for linha in range(2, len(sheet_selecionada['A']) + 1):
    nomeContato = sheet_selecionada['A%s' % linha].value
    mensagemContato = sheet_selecionada['B%s' % linha].value

    navegadorChrome.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]').send_keys(nomeContato)
    tempoEspera.sleep(3)

    teclasTeclado.press('enter')

    tempoEspera.sleep(3)

    navegadorChrome.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').send_keys(mensagemContato)

    tempoEspera.sleep(3)

    teclasTeclado.press('enter')


tempoEspera.sleep(25)