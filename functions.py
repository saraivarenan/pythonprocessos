import os, shutil
import fitz # install using: pip install PyMuPDF
from pdfminer.high_level import extract_text
from datetime import date
import PyPDF2 as PyPDF2
import pyautogui
import time
from os import path
from glob import glob
from zipfile import ZipFile
import tabula as tabula
import re



class downloadfile:
    def criarPasta(self):
        dir = os.path.join("C:\\","temp","python")
        if not os.path.exists(dir):
            os.mkdir(dir)
    def moveFile(cnpj):
        day = str(date.today().day)
        year = str(date.today().year)
        month = str(date.today().month)
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
    def clickImage(imagem):
        if  pyautogui.locateOnScreen(imagem, grayscale=True) != None:
            imageClick = pyautogui.locateOnScreen(imagem, grayscale=True);
            local = pyautogui.center(imageClick);
            pyautogui.click(local);
        else:
            pyautogui.hotkey('ctrl', 'r')
            time.sleep(3)
    def AlterarCnpjDowload(cnpj,imagem, alerta, check, btnGuia, alterarPerfil,fecharDownload):
        pyautogui.click(1100, 210);
       # downloadfile.clickImage(alterarPerfil)
        time.sleep(3)
        downloadfile.clickImage(imagem)
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
        pyautogui.press('pagedown')
        time.sleep(2);
        downloadfile.clickImage(check)
        time.sleep(3);
        downloadfile.clickImage(btnGuia);
        time.sleep(4);
        pyautogui.press('enter');
        time.sleep(1)
        downloadfile.clickImage(fecharDownload)
        downloadfile.moveFile(cnpj)
    def find_ext(dr, ext):
        return glob(path.join(dr, "*.{}".format(ext)))

    def getValorPdf(file):
        text = extract_text(file)
        parsed = ''.join(text)
        last = parsed.split()
        idvalor = last.index('Valor:')
        print(last[idvalor + 2])
        return last[idvalor + 2]

    def retornarCnpj(file):
        with fitz.open(file) as doc:
            text = ""
            for page in doc:
                text += page.get_text()
        lista = text.split("\n")
        return lista[40]

    def retornaVencimento(file):
        with fitz.open(file) as doc:
            text = ""
            for page in doc:
                text += page.get_text()
        lista = text.split("\n")
        return lista[48]

    def retonarValor(file):
        with fitz.open(file) as doc:
            text = ""
            for page in doc:
                text += page.get_text()
        lista = text.split("\n")
        return lista[52]



