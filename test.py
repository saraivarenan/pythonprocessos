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


def pesquisarFilePdf():
    filepathpdf = "07_2017"
    cnpj = "11004650000121"
    pathfolder = "C:\\Users\\lucas.gomes\\Documents\\" + filepathpdf + "\\" + cnpj
    pathfile = "".join(map(str, downloadfile.find_ext(pathfolder, "pdf")))
    print(pathfile)

if __name__ == "__main__":
    pesquisarFilePdf()
