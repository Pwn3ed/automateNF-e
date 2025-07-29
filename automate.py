from sys import executable
import time
import dotenv
import unidecode
from keyboard import wait
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from dotenv import load_dotenv
import os


def data_read():
    data = pd.read_excel('boleto.xlsx')
    df = pd.DataFrame(data)

    load_dotenv()
    username = os.getenv("LOGIN")
    password = os.getenv("PASSWORD")
    path = os.getenv("PATH_URL")

    return df, username, password, path


def open_browser():
    service = Service(executable_path='./geckodriver')
    options = webdriver.FirefoxOptions()
    firefox_profile = webdriver.FirefoxProfile()
    options.profile = firefox_profile

    browser = webdriver.Firefox(service=service, options=options)

    return browser


def login(browser, username, password, path):
    browser.get(path + '/issweb/paginas/login')

    browser.find_element(By.XPATH, '//*[@id="username"]').send_keys(username)
    browser.find_element(By.XPATH, '//*[@id="password"]').send_keys(password)
    browser.find_element(By.XPATH, '//*[@id="j_idt110"]').click()

    try:
        element = WebDriverWait(browser, 10).until(EC.url_changes(path + '/issweb/paginas/login'))
        browser.get(path + '/issweb/paginas/admin/notafiscal/convencional/emissaopadrao')
    except:
        browser.get(path + '/issweb/paginas/admin/notafiscal/convencional/emissaopadrao')


def change_data(browser, df, i):
    data_change = '31052025'  # CHANGE DD/MM/YYYY
    browser.find_element(By.XPATH, '//*[@id="formEmissaoNFConvencional:imDataCompetencia_input"]').click()
    browser.find_element(By.XPATH, '//*[@id="formEmissaoNFConvencional:imDataCompetencia_input"]').send_keys(data_change + Keys.TAB)
    time.sleep(1)


def find_cpf_cpnj(browser, df, i):
    n = len(df.loc[i]["CPF"].replace('.', '').replace('/','').replace('-',''))
    tipo_select_id = 'formEmissaoNFConvencional:tipoPessoa_label'
    tipo_fisica_id = 'formEmissaoNFConvencional:tipoPessoa_0'
    tipo_juridica_id = 'formEmissaoNFConvencional:tipoPessoa_1'


    cpf_id = 'formEmissaoNFConvencional:itCpf'
    cnpj_id = 'formEmissaoNFConvencional:itCnpj'
    print(f'i = {i}')

    if n == 11:  # CPF
        if browser.find_element(By.ID, tipo_select_id).text == 'Jurídica':
            try:
                browser.find_element(By.ID, tipo_select_id).click()
            except:
                print('error')
                exit()
            try:
                browser.find_element(By.ID, tipo_fisica_id).click()
            except:
                exit()

        wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        try:
            browser.find_element(By.ID, cpf_id).click()
        except:
            print('error')
            exit()
        try:
            browser.find_element(By.ID, cpf_id).send_keys(df.loc[i]["CPF"] + Keys.TAB)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()

    if n == 14:  # CNPJ
        if browser.find_element(By.ID, tipo_select_id).text == 'Física':
            try:
                browser.find_element(By.ID, tipo_select_id).click()
            except:
                print('error')
                exit()

            try:
                browser.find_element(By.ID, tipo_juridica_id).click()
            except:
                print('error')
                exit()

        wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        try:
            browser.find_element(By.ID, cnpj_id).click()
        except:
            print('error')
            exit()

        try:
            browser.find_element(By.ID, cnpj_id).send_keys(df.loc[i]["CPF"] + Keys.TAB)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()

    print(f'name = {df.loc[i]["NOME"]}')


def compare_names(browser, df, i):
    nome_id = 'formEmissaoNFConvencional:razaoNome'
    web_name = unidecode.unidecode(browser.find_element(By.ID, nome_id).get_attribute('value')).upper()
    excel_name = unidecode.unidecode(df.loc[i]["NOME"]).upper()
    if web_name != excel_name:
        try:
            browser.find_element(By.ID, nome_id).send_keys((Keys.CONTROL + 'a') + Keys.BACK_SPACE)
        except:
            print('error')
            exit()
        try:
            browser.find_element(By.ID, nome_id).send_keys(excel_name)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()


def fill_services(browser, df, i):
    descricao_id = 'formEmissaoNFConvencional:descricaoItem'
    if browser.find_element(By.ID, descricao_id) != 'MONITORAMENTO':
        try:
            browser.find_element(By.ID, descricao_id).send_keys('')
            time.sleep(0.5)
            # time.sleep(2)
            browser.find_element(By.ID, descricao_id).send_keys('monitoramento'.upper())
        except:
            print('error')
            exit()


def fill_services_value(browser, df, i):
    valor_id = 'formEmissaoNFConvencional:vlrUnitario_input'
    num = str(df.loc[i]["VALOR"]).split('.')

    if browser.find_element(By.ID, valor_id) is not num:
        try:
            browser.find_element(By.ID, valor_id).send_keys((Keys.CONTROL + 'a') + Keys.BACK_SPACE)
            browser.find_element(By.ID, valor_id).send_keys(num[0] + ',' + num[1] + (Keys.TAB * 3))
            print(f'value = {num[0]},{num[1]}')
        except:
            print('error fill_services_value()')
            exit()


def add_service(browser, df, i):
    add_btn_xpath = '/html/body/section/div/section/form/div[2]/div[1]/section/div/div[4]/div[2]/div[1]/div[2]/div[5]/button/span[2]'
    service_added = '//*[@id="formEmissaoNFConvencional:listaItensNota_data"]/tr/td[1]'
    try:
        time.sleep(0.5)
        browser.find_element(By.XPATH, add_btn_xpath).click()
        # wait = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, service_added)))
        print('service added')
    except :
        print('error add_service')
        exit()


def finish(browser, df, i):
    page_confirm_btn_1 = '//*[@id="frmActions:btnDefault"]/span[2]'
    page_confirm_btn_2 = '//*[@id="frmActions:j_idt480"]/span[2]'

# OLD VERSION DEPRECEATED
    # try:
    #     browser.execute_script('scrollBy(0,-10000)')
    #     print('scrolled up')
    # except:
    #     print('error')
    #     exit()
    try:
        time.sleep(0.5)
        # time.sleep(2)
        browser.find_element(By.XPATH, page_confirm_btn_1).click()  # Page confirm
    except:
        print('error1')
        exit()
    try:
        time.sleep(0.5)
        browser.find_element(By.XPATH, page_confirm_btn_2).click()  # Page confirm of the page confirm
        # time.sleep(5)
        print('confirm button')
    except:
        print('error2')
        exit()


def comparison(browser, df, i):
    total_value_xpath = '//*[@id="formEmissaoNFConvencional:valoresNota"]/div/div/div[3]/div'
    total_value_element = browser.find_element(By.XPATH, total_value_xpath)
    try:
        wait = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element( total_value_element, str(df.loc[i]["VALOR"]) + '0' ))

        time.sleep(3)

        total = float(browser.find_element(By.XPATH, total_value_xpath).text.replace(',', '.'))
        print(f'comparison: total = {total} = {df.loc[i]["VALOR"]} = df.loc[i]["VALOR"]')
        if total != df.loc[i]["VALOR"]:
            print('ERROR TO MATCH VALUES')
            return False
    except:
        print('error')
        exit()

    print('matched values')
    return True


def new_nfe(browser, df, i):
    # (feat) SEND AUTOMATIC EMAILS COMPARING THE BOOLEAN 'TRUE' OR 'FALSE'
    new_nfe_btn = '//*[@id="formEmissaoNFConvencional:j_idt783"]/span[2]'

    try:
        wait = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, new_nfe_btn)))
        browser.find_element(By.XPATH, new_nfe_btn).click()
        print('new nfe')
    except:
        print('error')
        exit()


def test():
    print('starting main ...')
    df, username, password, path = data_read()
    browser = open_browser()
    login(browser, username, password, path)


def main():
    print('starting main ...')
    df, username, password, path = data_read()
    browser = open_browser()
    login(browser, username, password, path)
    start_line = 3  # default = 3
    # start_line = start_line + 2  # To start one number plus "i" && +2 to continue case the for loop breaks
    end_line = 6
    # for i in range(start_line - 2, end_line - 1):
    for i in range(start_line - 2, end_line - 1):  # Default: for i in range(0, len(df)) or range((2)-2, len(df)):
        print(f'before call function in loop {time.perf_counter()}')
        # change_data(browser, df, i)  # optional - CONFIRM THIS, DEFAULT IS COMMENTED
        find_cpf_cpnj(browser, df, i)
        compare_names(browser, df, i)
        fill_services(browser, df, i)
        fill_services_value(browser, df, i)
        add_service(browser, df, i)
        # if comparison(browser, df, i):
        #     finish(browser, df, i)
        finish(browser, df, i)
        new_nfe(browser, df, i)
        print(f'after call function in loop {time.perf_counter()}')


if __name__ == '__main__':
    print('Starting code ...')
    main()
    time.sleep(1)
    print('\nCompleted all tasks with success.')


