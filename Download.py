import os
import time
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
consulta = 'http://gomnet.ampla.com/ConsultaObra.aspx'
username = login['A1'].value
password = login['A2'].value

# Configurações do browser
profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', os.getcwd())
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.ms-excel')
driver = webdriver.Firefox(profile)

if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    driver.get(consulta)

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_NumSOB')
    with open('sobs.txt') as data:
        datalines = (line.rstrip('\r\n') for line in data)
        for line in datalines:
            window_before = driver.window_handles[0]
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_NumSOB').clear()
            sob = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_TextBox_NumSOB')))
            sob.send_keys(line)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_ImageButton_Enviar'))).click()
            try:
                # Busca pela coluna com o número da Sob
                numSob = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[4]/td/div[3]/table/tbody/tr[2]/td[8][contains(text(), "' + line + '")]')
                if numSob.is_displayed():
                    numSobArquivo = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Gridview_GomNet1"]/tbody/tr[2]/td[8]').text
                    numTrabArquivo = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Gridview_GomNet1"]/tbody/tr[2]/td[4]').text
                    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Gridview_GomNet1_ctl02_ImageButton_OrcamentoConstrutivo"]').click()
                    window_after = driver.window_handles[1]
                    driver.switch_to_window(window_after)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="GridView_Solicitacoes_ctl02_ImageButton_Excel"]'))).click()
                    driver.close()
                    driver.switch_to_window(window_before)
                    # Aguarda o download completo do arquivo para então renomeá-lo
                    while os.path.exists('Relatorio_Gestao_Obra.xls.part'):
                        time.sleep(5)
                    if os.path.isfile('Relatorio_Gestao_Obra.xls'):
                        os.rename('Relatorio_Gestao_Obra.xls', numSobArquivo + ' ' + numTrabArquivo + '.xls')
            except NoSuchElementException:
                log = open("ErroSobs.txt", "a")
                log.write(line + "\n")
                log.close()
                continue
