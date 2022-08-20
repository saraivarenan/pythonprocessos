import pandas as pd
import os, shutil
from datetime import date
import PyPDF2 as pypdf2
import pyautogui
import time
from os import path
from glob import glob
from zipfile import ZipFile

class EcacGuia:
    def moveFile(cnpj):

        day = str(date.today().day)
        year = str(date.today().year)
        month = str(date.today().month -1)
        pastadownload = str(month+"_"+ year)
        dir = os.path.join("C:\\Users\\lucas.gomes\\Documents",pastadownload, cnpj)
        print(dir)
        if not os.path.exists(dir):
            os.makedirs(dir)
        path = r"C:/Users/lucas.gomes/Downloads/"
        moveto = dir+"\\"
        files = os.listdir(path)
        files.sort()
        for f in files:
            src = path+f
            dst = moveto+f
            shutil.move(src,dst)
    def find_ext(dr, ext):
        return glob(path.join(dr, "*.{}".format(ext)))
    def excluirArquivo(pathFile):
        path = r"C:/Users/lucas.gomes/Downloads/"
        files = os.listdir(path)
        print(len(files))
        if (len(files) >0):
            for f in files:
                os.remove(path+f)
                print("Arquivo excluido"+f)
    def extrairArquivo(cnpj, dir):
        print("extraindo arquivo zip para .pdf. Cliente: ", cnpj)
        path1 = "".join(map(str, EcacGuia.find_ext(dir, "zip")))
        with ZipFile(path1, 'r') as zipObj:
            # Extract all the contents of zip file in different directory
            zipObj.extractall(dir)
            print('Arquivo extraido para pdf. Cliente: ', cnpj)
    def clickImage(imagem):
      imageClick = pyautogui.locateOnScreen(imagem, grayscale=True);
      local = pyautogui.center(imageClick);
      pyautogui.click(local)


    def validarSemMovimenta(check,btnGuia,fechardown, cnpj, logexecucao):
        year = str(date.today().year)
        month = str(date.today().month - 1)
        pastadownload = str(month + "_" + year)
        dir = os.path.join("C:\\Users\\lucas.gomes\\Documents", pastadownload, cnpj)
        print(dir)
        semMovimenta = r"images/clickSemMovimento.png"
        if pyautogui.locateOnScreen(check, grayscale=True) != None:
            EcacGuia.clickImage(check)
            time.sleep(10);
            EcacGuia.clickImage(btnGuia);
            time.sleep(4);
            pyautogui.press('enter');
            time.sleep(1)
            EcacGuia.moveFile(cnpj)
            EcacGuia.extrairArquivo(cnpj, dir)
        else:
            EcacGuia.clickImage(semMovimenta)
            time.sleep(2)
            pyautogui.click(1100, 500);
            EcacGuia.moveFile(cnpj)

def bot():
    imagem = r"images\inserirCnpj.png";
    alerta = r"images\alertarobo.png";
    check = r"images\checkbox.png";
    btnAlterarPerfil = r"images\iconeAlterar.png";
    linkAssinar = r"images\assinarTransmitir.png";
    btnGuia = r"images\guia.png";
    btndemonstrativo = r"images\btnDemo.png";
    btnGov = r"images\btnEntrarGov.png";
    linkCert = r"images\linkCert.png";
    btnOkCert = r"images\btnOKCert.png";
    fecharDownload = r"images\fecharDownload.png";
    fecharNavagador = r"images\fecharNavegador.png";
    fechardown = r"images\imagemFechar.png";
    btnRecarregar = r"images\btnRecarregar.png";
    file = r"dadosUm.csv";
    df = pd.read_csv(file)
    year = str(date.today().year)
    month = str(date.today().month-1)
    listExecLog = []
    logexecucao = "C:/Users/lucas.gomes/Documents/logExecucao_" + month + "_" + year + ".txt"
    listExecLog.append('{}{}{}\n'.format("CNPJ", ",", "STATUS"))
    indice = 0
    indiceEntrada = len(df)
    while indice < indiceEntrada:
        try:

            print(indice)
         # fileTxt = open( logexecucao,"w+")
            cnpj = df.get("cpf")[indice]
            cnpji = cnpj.replace(".","")
            cnpji = cnpji.replace("/","")
            cnpji = cnpji.replace("-","")
         # # realizado com a o tamanho Size(width=1366, height=768)
            pyautogui.press('winleft')
            time.sleep(2)
            pyautogui.write("chrome")
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(5)
            EcacGuia.clickImage(btnGov)
            time.sleep(4)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')
            #EcacGuia.clickImage(linkCert)
            time.sleep(2)
            EcacGuia.clickImage(btnOkCert)
            time.sleep(5)
            EcacGuia.clickImage(btndemonstrativo)
            time.sleep(3);
            EcacGuia.clickImage(linkAssinar)
            time.sleep(3);
            pyautogui.click(1100, 210);

            time.sleep(3)
            EcacGuia.clickImage(imagem)
            time.sleep(2)
            pyautogui.typewrite(cnpji, interval=0.25);
            time.sleep(4);
            pyautogui.press("tab");
            time.sleep(2)
            pyautogui.press("enter");

            time.sleep(5)
             # pyautogui.press('pagedown')
            time.sleep(2);
            EcacGuia.validarSemMovimenta(check, btnGuia, fechardown, cnpji, logexecucao)

            listExecLog.append('{}{}{}\n'.format(cnpji, ",", "OK"))
            with open(logexecucao, "w") as logexec:
                logexec.writelines(listExecLog)

        except Exception as erro:
            print(" Ocorreu um erro", erro)
            time.sleep(3)
            listExecLog.append('{}{}{}\n'.format(cnpji, ",", "NOK"))
            with open(logexecucao, "w") as logexec:
                logexec.writelines(listExecLog)

        finally:
            EcacGuia.clickImage(fecharNavagador)
            EcacGuia.excluirArquivo("Verificar se ainda existe arquivo")


        indiceSaida =indice+1
        indice =indiceSaida

if __name__ == "__main__":
    bot()