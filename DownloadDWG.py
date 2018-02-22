import os
import time
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import openpyxl

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
url_dwnld = 'http://gomnet.ampla.com/Upload.aspx?numsob='
username = login['A1'].value
password = login['A2'].value

# Configurações do Browser
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": os.getcwd(),
         "download.prompt_for_download": False}
chromeOptions.add_experimental_option("prefs", prefs)
# chromeOptions.add_argument('--headless')
driver = webdriver.Chrome(chrome_options=chromeOptions)

if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    with open('sobs.txt') as data:
        datalines = (line.rstrip('\r\n') for line in data)
        for line in datalines:
            driver.get(url_dwnld + line)
            window_before = driver.window_handles[0]
            try:
                # Busca pela versão mais atual do dwg da sob
                try:
                    revObra9 = driver.find_element_by_partial_link_text('REVOBRA_09.')
                    if revObra9.is_displayed():
                        revObra9.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra8 = driver.find_element_by_partial_link_text('REVOBRA_08')
                    if revObra8.is_displayed():
                        revObra8.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra7 = driver.find_element_by_partial_link_text('REVOBRA_07')
                    if revObra7.is_displayed():
                        revObra7.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra6 = driver.find_element_by_partial_link_text('REVOBRA_06')
                    if revObra6.is_displayed():
                        revObra6.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra5 = driver.find_element_by_partial_link_text('REVOBRA_05')
                    if revObra5.is_displayed():
                        revObra5.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra4 = driver.find_element_by_partial_link_text('REVOBRA_04')
                    if revObra4.is_displayed():
                        revObra4.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra3 = driver.find_element_by_partial_link_text('REVOBRA_03')
                    if revObra3.is_displayed():
                        revObra3.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra2 = driver.find_element_by_partial_link_text('REVOBRA_02')
                    if revObra2.is_displayed():
                        revObra2.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra1 = driver.find_element_by_partial_link_text('REVOBRA_01')
                    if revObra1.is_displayed():
                        revObra1.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra = driver.find_element_by_partial_link_text('REVOBRA')
                    if revObra.is_displayed():
                        revObra.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev9 = driver.find_element_by_partial_link_text('REV09')
                    if rev9.is_displayed():
                        rev9.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev8 = driver.find_element_by_partial_link_text('REV08')
                    if rev8.is_displayed():
                        rev8.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev7 = driver.find_element_by_partial_link_text('REV07')
                    if rev7.is_displayed():
                        rev7.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev6 = driver.find_element_by_partial_link_text('REV06')
                    if rev6.is_displayed():
                        rev6.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev5 = driver.find_element_by_partial_link_text('REV05')
                    if rev5.is_displayed():
                        rev5.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev4 = driver.find_element_by_partial_link_text('REV04')
                    if rev4.is_displayed():
                        rev4.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev3 = driver.find_element_by_partial_link_text('REV03')
                    if rev3.is_displayed():
                        rev3.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev2 = driver.find_element_by_partial_link_text('REV02')
                    if rev2.is_displayed():
                        rev2.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev1 = driver.find_element_by_partial_link_text('REV01')
                    if rev1.is_displayed():
                        rev1.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
                try:
                    dwg = driver.find_element_by_partial_link_text('dwg')
                    if dwg.is_displayed():
                        dwg.click()
                        window_after = driver.window_handles[1]
                        driver.switch_to_window(window_after)
                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(5)
                        continue
                except NoSuchElementException:
                    pass
            except NoSuchElementException:
                log = open("ErroSobs.txt", "a")
                log.write(line + "\n")
                log.close()
                continue
