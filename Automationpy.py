import pandas as pd
import os, shutil
from datetime import date
import PyPDF2 as pypdf2
import pyautogui
import time
from os import path
from glob import glob
from zipfile import ZipFile
class downloadfile:
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
    file = r"dadosUm.csv";
    df = pd.read_csv(file)
    year = str(date.today().year)
    month = str(date.today().month)

    def interacaoDownload(cnpji,imagem,alerta,check,btnGuia,btnAlterarPerfil,fechardown,df,listExecLog,logexecucao,indice):
        try:
            indiceEntrada = len(df)
            while  indice < indiceEntrada:
        #for indice, data in df.iterrows():
                print(indice)
                cnpj = df.get("cpf")[indice]
                cnpji = cnpj.replace(".", "")
                cnpji = cnpji.replace("/", "")
                cnpji = cnpji.replace("-", "")
                print(cnpji)
                downloadfile.AlterarCnpjDowload(cnpji,imagem,alerta,check,btnGuia,btnAlterarPerfil,fechardown,logexecucao)
                listExecLog.append('{}{}{}\n'.format(cnpji,",","OK"))
                with open(logexecucao, "w") as logexec:
                    logexec.writelines(listExecLog)
                indice += 1
            return indice
        except:
            listExecLog.append('{}{}{}\n'.format(cnpji, ",", "NOK"))
            with open(logexecucao, "w") as logexec:
                logexec.writelines(listExecLog)
            pyautogui.hotkey('ctrl', 'r')
            return indice


    def criarPasta(self):
        dir = os.path.join("C:\\","temp","python")
        if not os.path.exists(dir):
            os.mkdir(dir)
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
        print("extraindo arquivo zip para .pdf. Cliente: ",cnpj)
        path1 = "".join(map(str, downloadfile.find_ext(dir, "zip")))
        with ZipFile(path1, 'r') as zipObj:
            # Extract all the contents of zip file in different directory
            zipObj.extractall(dir)
            print('Arquivo extraido para pdf. Cliente: ', cnpj)
    def clickImage(imagem, cnpji, logexecucao):
        listExecLog = []

        if (pyautogui.locateOnScreen(imagem, grayscale=True) != None):
            imageClick = pyautogui.locateOnScreen(imagem, grayscale=True);
            local = pyautogui.center(imageClick);
            pyautogui.click(local)
            time.sleep(3)
        else :
            listExecLog.append('{}{}{}\n'.format(cnpji, ",", "NOK"))
            with open(logexecucao, "w") as logexec:
                logexec.writelines(listExecLog)

    def AlterarCnpjDowload(cnpj,imagem, alerta, check, btnGuia, alterarPerfil,fechardown, logexecucao):
        pyautogui.click(1100, 210);
       # downloadfile.clickImage(alterarPerfil)
        time.sleep(3)
        downloadfile.clickImage(imagem, cnpj, logexecucao)
        time.sleep(2)
        pyautogui.typewrite(cnpj, interval=0.25);
        time.sleep(4);
        pyautogui.press("tab");
        time.sleep(2)
        pyautogui.press("enter");
        for i in range(0, 9):
            if pyautogui.locateOnScreen(alerta, grayscale=True) != None:
                pyautogui.rightClick(780, 382);
                time.sleep(4)
            else:
                pyautogui.PAUSE=0.5;
                break
        time.sleep(5)
        #pyautogui.press('pagedown')
        time.sleep(2);
        downloadfile.clickImage(check,cnpj, logexecucao)
        time.sleep(3);
        downloadfile.clickImage(btnGuia, cnpj, logexecucao);
        time.sleep(4);
        pyautogui.press('enter');
        time.sleep(1)
        downloadfile.clickImage(fechardown, cnpj, logexecucao)
        downloadfile.moveFile(cnpj)
    def find_ext(dr, ext):
        return glob(path.join(dr, "*.{}".format(ext)))

    def getValorPdf(file):
        pdf_file = open(file, 'rb')
        read_pdf = pypdf2.PdfFileReader(pdf_file)
        page = read_pdf.getPage(0)
        page_content = page.extract_text()
        parsed = ''.join(page_content)
        last = parsed.split()
        idvalor = last.index('Valor:')
        return last[idvalor + 1]
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
    file = r"dadosUm.csv";
    df = pd.read_csv(file)
    year = str(date.today().year)
    month = str(date.today().month)
    listExecLog = []
    logexecucao = "C:/Users/lucas.gomes/Documents/logExecucao_" + month + "_" + year + ".txt"
    try:
       # fileTxt = open( logexecucao,"w+")
        cnpj = df.get("cpf")[0]
        cnpji = cnpj.replace(".","")
        cnpji = cnpji.replace("/","")
        cnpji = cnpji.replace("-","")
        # realizado com a o tamanho Size(width=1366, height=768)
        pyautogui.press('winleft')
        time.sleep(2)
        pyautogui.write("chrome")
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(5)
        downloadfile.clickImage(btnGov,"","")
        time.sleep(2)
        downloadfile.clickImage(linkCert,"","")
        time.sleep(2)
        downloadfile.clickImage(btnOkCert,"","")
        time.sleep(5)
        downloadfile.clickImage(btndemonstrativo,"","")
        time.sleep(3);
        downloadfile.clickImage(linkAssinar,"","")
        time.sleep(3);
        listExecLog.append('{}{}{}\n'.format("CNPJ",",","STATUS"))
        indice = 0
        indiceEntrada =len(df)
        while indice <= indiceEntrada:
            indiceinteracao = 1+ downloadfile.interacaoDownload(cnpji,imagem,alerta,check,btnGuia,btnAlterarPerfil,fechardown,df,listExecLog,logexecucao, indice)
            indice =indiceinteracao
    except Exception as erro:
        print(" Ocorreu um erro", erro)
        time.sleep(3)


    finally:
        downloadfile.clickImage(fecharNavagador,cnpji,logexecucao)



if __name__ == "__main__":
    bot()