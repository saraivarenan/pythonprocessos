import pandas as pd
import os, shutil
from datetime import date
import PyPDF2 as pypdf2
import pyautogui
import time
from os import path
from glob import glob
from zipfile import ZipFile

try:
    file = r"dadosUm.csv";
    df = pd.read_csv(file)
    listExecLog = []
    logexecucao = "C:/Users/lucas.gomes/Documents/logProcurarArquivo_07_2022.txt"
    for indice, data in df.iterrows():
        cnpj = df.get("cpf")[indice]
        cnpji = cnpj.replace(".", "")
        cnpji = cnpji.replace("/", "")
        cnpji = cnpji.replace("-", "")
        dir = os.path.join("C:\\Users\\lucas.gomes\\Documents\\backup_07", cnpji)
        contemarquivo = glob(path.join(dir, "*.{}".format("pdf")))
        if len(contemarquivo) ==1:
            listExecLog.append('{}{}{}\n'.format(cnpji, ",", "OK"))
            with open(logexecucao, "w") as logexec:
                 logexec.writelines(listExecLog)
            print("1")
        else:
            # listExecLog.append('{}{}{}\n'.format(cnpji, ",", "NOK"))
            # with open(logexecucao, "w") as logexec:
            #     logexec.writelines(listExecLog)
            print("0")


except:
    print("errou")