import time
import unidecode
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


def data_read():
    data = pd.read_excel('boleto.xlsx')
    df = pd.DataFrame(data, columns = ['CPF', 'VALOR', 'NOME', 'BOOLEAN'])
    login_data = pd.read_excel('login.xlsx')
    return df, login_data


def open_browser():
    browser = webdriver.Chrome('./chromedriver')
    return browser


def login(browser, login_data):
    # old_version = browser = webbrowser.Chrome()
    browser.get('http://45.175.63.178:8080/issweb/paginas/login')
    browser.find_element_by_xpath('//*[@id="username"]').send_keys(login_data.loc[0][0])
    browser.find_element_by_xpath('//*[@id="password"]').send_keys(login_data.loc[1][0])
    browser.find_element_by_xpath('//*[@id="j_idt110"]').click()

    try:
        element = WebDriverWait(browser, 10).until(EC.url_changes('http://45.175.63.178:8080/issweb/paginas/login'))
        browser.get('http://45.175.63.178:8080/issweb/paginas/admin/notafiscal/convencional/emissaopadrao')
    except:
        browser.get('http://45.175.63.178:8080/issweb/paginas/admin/notafiscal/convencional/emissaopadrao')


# C H A N G E  D A T A
def change_data(browser, df, i):
     data_change = '03062023'
     browser.find_element_by_xpath('//*[@id="formEmissaoNFConvencional:imDataCompetencia_input"]').click()
     browser.find_element_by_xpath('//*[@id="formEmissaoNFConvencional:imDataCompetencia_input"]').send_keys(data_change + Keys.TAB)
    

def find_cpf_cpnj(browser, df, i):
    n = len(df.loc[i][0].replace('.', '').replace('/','').replace('-',''))
    tipo_select_id = 'formEmissaoNFConvencional:tipoPessoa_label'
    tipo_fisica_id = 'formEmissaoNFConvencional:tipoPessoa_0'
    tipo_juridica_id = 'formEmissaoNFConvencional:tipoPessoa_1'

    cpf_id = 'formEmissaoNFConvencional:itCpf'
    cnpj_id = 'formEmissaoNFConvencional:itCnpj'
    print(f'i = {i}')


    if n == 11: # CPF
        if browser.find_element_by_id(tipo_select_id).text == 'Jurídica':
            try:
                browser.find_element_by_id(tipo_select_id).click()
            except:
                print('error')
                exit()
            try:
                browser.find_element_by_id(tipo_fisica_id).click()
            except:
                exit()

        wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        try:
            browser.find_element_by_id(cpf_id).click()
        except:
            print('error')
            exit()
        try:
            browser.find_element_by_id(cpf_id).send_keys(df.loc[i][0] + Keys.TAB)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()

    if n == 14: # CNPJ
        if browser.find_element_by_id(tipo_select_id).text == 'Física':
            try:
                browser.find_element_by_id(tipo_select_id).click()
            except:
                print('error')
                exit()

            try:
                browser.find_element_by_id(tipo_juridica_id).click()
            except:
                print('error')
                exit()

        wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        try:
            browser.find_element_by_id(cnpj_id).click()
        except:
            print('error')
            exit()

        try:
            browser.find_element_by_id(cnpj_id).send_keys(df.loc[i][0] + Keys.TAB)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()

    print(f'name = {df.loc[i][2]}')


def compare_names(browser, df, i):
    # Need to compare the Web_name with the Excel_name
    nome_id = 'formEmissaoNFConvencional:razaoNome'
    web_name = unidecode.unidecode(browser.find_element_by_id(nome_id).get_attribute('value')).upper()
    excel_name = unidecode.unidecode(df.loc[i][2]).upper()
    if web_name != excel_name: # compare
        try:
            browser.find_element_by_id(nome_id).send_keys((Keys.CONTROL + 'a') + Keys.BACK_SPACE)
        except:
            print('error')
            exit()
        try:
            browser.find_element_by_id(nome_id).send_keys(excel_name)
            wait = WebDriverWait(browser, 10).until(EC.invisibility_of_element_located((By.ID, 'dStatus')))
        except:
            print('error')
            exit()


def fill_services(browser, df, i):
    descricao_id = 'formEmissaoNFConvencional:descricaoItem'
    if browser.find_element_by_id(descricao_id) != 'MONITORAMENTO':
        try:
            browser.find_element_by_id(descricao_id).send_keys('')
            time.sleep(0.5)
            # time.sleep(2)
            browser.find_element_by_id(descricao_id).send_keys('monitoramento'.upper())
        except:
            print('error')
            exit()


def fill_services_value(browser, df, i):
    valor_id = 'formEmissaoNFConvencional:vlrUnitario_input'
    num = str(df.loc[i][1]).split('.')

    if browser.find_element_by_id(valor_id) is not num:
        try:
            browser.find_element_by_id(valor_id).send_keys((Keys.CONTROL + 'a') + Keys.BACK_SPACE)
            browser.find_element_by_id(valor_id).send_keys(num[0] + ',' + num[1] + Keys.TAB)
            print(f'value = {num[0]},{num[1]}')
        except:
            print('error fill_services_value()')
            exit()


def add_service(browser, df, i):
    add_btn_xpath = '//*[@id="formEmissaoNFConvencional:btnAddItem"]/span[1]'
    service_added = '//*[@id="formEmissaoNFConvencional:listaItensNota_data"]/tr/td[1]'
    try:
        browser.find_element_by_xpath(add_btn_xpath).click()
        # time.sleep(2)
        wait = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, service_added)))
        print('service added')
    except :
        print('error')
        exit()


def finish(browser, df, i):
    page_confirm_btn_1 = '//*[@id="frmActions:btnDefault"]/span[2]'
    page_confirm_btn_2 = '//*[@id="frmActions:j_idt479"]/span[2]'

    try:
        browser.execute_script('scrollBy(0,-10000)')
        print('scrolled up')
    except:
        print('error')
        exit()
    try:
        time.sleep(0.5)
        browser.find_element_by_xpath(page_confirm_btn_1).click()  # Page confirm
        print('confirm first button')
    except:
        print('error')
        exit()
    try:
        time.sleep(0.5)
        browser.find_element_by_xpath(page_confirm_btn_2).click()  # Page confirm of the page confirm
        print('confirm second button')
    except:
        print('error')
        exit()


def comparison(browser, df, i):
    total_value_xpath = '//*[@id="formEmissaoNFConvencional:valoresNota"]/div/div/div[3]/div'
    total_value_element = browser.find_element_by_xpath(total_value_xpath)
    try:
        wait = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element( total_value_element, str(df.loc[i][1]) + '0' ))
        
        time.sleep(3)

        total = float(browser.find_element_by_xpath(total_value_xpath).text.replace(',', '.'))
        print(f'comparison: total = {total} = {df.loc[i][1]} = df.loc[i][1]')
        if total != df.loc[i][1]:
            print('ERROR TO MATCH VALUES')
            return False
    except:
        print('error')
        exit()

    print('matched values')
    return True


def new_nfe(browser, df, i):
    # NEED TO SEND AUTOMATIC EMAILS COMPARING THE BOOLEAN 'TRUE' OR 'FALSE'
    new_nfe_btn = '//*[@id="formEmissaoNFConvencional:j_idt774"]'

    try:
        wait = WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, new_nfe_btn)))
        browser.find_element_by_xpath(new_nfe_btn).click()
        print('new nfe')
    except:
        print('error')
        exit()


def test():
    print('starting main ...')
    df, login_data = data_read()
    browser = open_browser()
    login(browser, login_data)


def main():
    print('starting main ...')
    df, login_data = data_read()
    browser = open_browser()
    login(browser, login_data)
    start_line = 3
    # start_line = start_line + 2  # Para começar um nome após o "i" # +2 para continuar o codigo
    end_line = 287
    # for i in range(start_line - 2, end_line - 1):
    for i in range(start_line - 2, end_line - 1):  # Default: for i in range(0, len(df)) or range((2)-2, len(df)):
        print(f'before call function in loop {time.perf_counter()}')
        # change_data(browser, df, i) #opcional
        time.sleep(1)
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


print('starting code ...')
main()
time.sleep(10)
print('finish')
