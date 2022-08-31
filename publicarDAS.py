import os
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
from config import caminhoDriver, urlSciweb, usuario, senha
from functions import downloadfile
from datetime import date

from glob import glob


try:
    file = r"dadosUm.csv";
    df = pd.read_csv(file)
    chromedriver = caminhoDriver
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"profile.default_content_settings.popups": 0,
         "security.default_personal_cert": "Select Automatically",
         "accept_untrusted_certs": True,
                 "directory_upgrade": True}
    chromeOptions.add_experimental_option("prefs",prefs)
    chromeOptions.add_argument('--allow-running-insecure-content')
    chromeOptions.add_argument('--ignore-certificate-errors')
    chromeOptions.add_argument('--disable-blink-features=AutomationControlled')
    chromeOptions.add_argument('--ash-force-desktop')
    chromeOptions.add_argument('--ignore-ssl-errors')
    chromeOptions.add_argument('--window-size=1366x768')
    #chromeOptions.add_argument('headless')
    year = str(date.today().year)
    mesAtual = str(date.today().month-1)

    logexecucao = "C:/Users/lucas.gomes/Documents/logExecucao_DAS" + mesAtual + "_" + year + ".txt"

    listExecLog = []
    driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)
    driver.get(urlSciweb)
    driver.maximize_window()

#Fazer login
    driver.find_element(by=By.ID,value="usuario").send_keys(usuario)
    driver.find_element(by=By.ID,value="senha").send_keys(senha)

    driver.find_element(by=By.XPATH, value="//input[@value='Entrar']").click();
    time.sleep(2)

#clicar em report 24 hs
    reportElement = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, ".icon-sci-report-geral")))
    reportElement.click();
    time.sleep(2)

#Clicar em publicar
    publicarElement = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, ".icon-publicar")))
    publicarElement.click();
    time.sleep(2)

#habilitar search pesquisa

    monthreferencia = str(date.today().month)
    mesAtual = str((date.today().month)-1)
    mesvencimento = str((date.today().month))

    monthNumber = (date.today().month)
    filepathpdf = mesAtual+"_"+year
    dataVenc = "20"+mesAtual+year
    if(monthNumber <10):
        valorRefenrencia = "0"+mesAtual+year
        dataVenc = "20"+"0"+mesvencimento+ year
    else:
        dataVenc ="20"+mesAtual+year
    listExecLog.append('{}{}{}\n'.format("CNPJ", ",", "STATUS"))
    path = r"C:\Users\lucas.gomes\Documents\Relatiorio_DAS_2022"
    files = downloadfile.find_ext(path, "pdf")
    #files.sort()
    for f in files:
        nomepasta =str(f)
        valorCnpjpdf = downloadfile.getValorPDF(nomepasta,'CNPJ:',1)
        cnpj = valorCnpjpdf
        cnpji = cnpj.replace(".", "")
        cnpji = cnpji.replace("/", "")
        cnpji = cnpji.replace("-", "")

        valorVencimento = downloadfile.getValorPDF(nomepasta,'até',1)
        valorpdf = downloadfile.getValorPDF(nomepasta,'Valor:',2)

        print(" ",valorpdf)
        selectEmpresa = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "select2-empresaId-container")))
        selectEmpresa.click();
        time.sleep(2)
        driver.switch_to.active_element.send_keys(cnpji)
        time.sleep(4)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(2)

        # pesquisar relatorio
        selectReport = driver.find_element(By.ID, value="select2-relatorioId-container")
        selectReport.click()
        time.sleep(2)
        driver.switch_to.active_element.send_keys("0072")
        time.sleep(6)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        # Preencher datas e valor
        driver.find_element(By.ID, value="dataReferencia").send_keys(valorRefenrencia)
        time.sleep(2)
        driver.find_element(By.ID, value="dataVencimento").send_keys(valorVencimento)
        time.sleep(2)
        driver.find_element(By.ID, value="valor").send_keys(valorpdf)
        time.sleep(2)

        body = driver.find_element(By.CSS_SELECTOR, "body")
        body.send_keys(Keys.PAGE_DOWN)
        time.sleep(1)
        driver.find_element(By.ID, value="descricao").send_keys('DAS'+valorRefenrencia)
        time.sleep(1)

        #selecionar arquivo
        selectEmpresa = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "select2-tipoGuiaId-container")))
        selectEmpresa.click();
        time.sleep(2)
        driver.switch_to.active_element.send_keys("N")
        time.sleep(4)
        driver.switch_to.active_element.send_keys(Keys.ENTER)

        # anexar documento
        time.sleep(3)

        upload = driver.find_element(By.ID, value="arquivo")
        # upload.click()
        upload.send_keys(os.path.abspath(nomepasta))

        time.sleep(3)
        driver.find_element(By.XPATH, value="//button[contains(.,'Publicar')]").click()
        time.sleep(5)
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#corpo > #avisosFixed > .acerto")))

        mensagemSucesso = driver.find_element(By.CSS_SELECTOR, value="#corpo > #avisosFixed > .acerto").text
        mensagemValidar = "Arquivo publicado com sucesso."

        assert mensagemSucesso.__eq__(mensagemValidar)
        time.sleep(5)

        print("Arquivo publicado")
        listExecLog.append('{}{}{}\n'.format(valorCnpjpdf, ",", "OK"))
        with open(logexecucao, "w") as logexec:
            logexec.writelines(listExecLog)
        driver.refresh();
except Exception as erro:
    print(erro)
    driver.quit()
finally:
    driver.quit()