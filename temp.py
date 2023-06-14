import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import numpy as np
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import glob
import chromedriver_autoinstaller
import datetime

def ultimo_arquivo():

    diretorio = r"C:\Users\TI\Downloads"
    arquivos_csv = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('.csv')]

    # ordenar os arquivos pela data de modificação, do mais recente para o mais antigo
    arquivos_csv = sorted(arquivos_csv, key=lambda arquivo: os.path.getmtime(os.path.join(diretorio, arquivo)), reverse=True)

    # pegar o arquivo mais recente
    ultimo_arquivo = arquivos_csv[0]

    return ultimo_arquivo, diretorio

# chromedriver_autoinstaller.install()

scope = ['https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive"]
   
credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
client = gspread.authorize(credentials)
filename = 'service_account_cemag.json'
sa = gspread.service_account(filename)

sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
worksheet = 'BD - ESTOQUE'
worksheet2 = 'ALMOX CENTRAL'
worksheet3 = 'SALDO SERRA'
worksheet4 = 'CONTROLE - ESTOQUE'

sh1 = sa.open(sheet)
wks1 = sh1.worksheet(worksheet)
wks2 = sh1.worksheet(worksheet2)
wks3 = sh1.worksheet(worksheet3)
wks4 = sh1.worksheet(worksheet4)

df = wks1.get()
df2 = wks2.get()
df3 = wks3.get()
df4 = wks4.get()

tabela = pd.DataFrame(df)
tabela2 = pd.DataFrame(df2)
tabela3 = pd.DataFrame(df3)
tabela4 = pd.DataFrame(df4)

cabecalho = wks1.row_values(1)
cabecalho2 = wks2.row_values(1)
cabecalho3 = wks3.row_values(1)
cabecalho4 = wks4.row_values(2)

def iframes(nav):

    iframe_list = nav.find_elements(By.CLASS_NAME,'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try: 
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

def saida_iframe(nav):
    nav.switch_to.default_content()

def listar(nav, classe):
    
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)


nav = webdriver.Chrome(r'C:\\Users\\TI\\anaconda3\\lib\\site-packages\\chromedriver_autoinstaller\\114\\chromedriver.exe')
time.sleep(1)
nav.maximize_window()
time.sleep(1)
nav.get('http://192.168.3.141/sistema')

nav.find_element(By.ID, 'username').send_keys('ti.dev') #ti.dev ti.cemag
time.sleep(2)
nav.find_element(By.ID, 'password').send_keys('cem@#161010')
time.sleep(1)
nav.find_element(By.ID, 'submit-login').click() 
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'bt_1892603865')))
time.sleep(1)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(3)


lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)['index'][0]
lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Saldos de recursos'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(6)

iframes(nav)
data_base = nav.find_element(By.XPATH,'/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
data_base.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
data_base.send_keys('h')
time.sleep(1.5)
data_base.send_keys(Keys.ENTER)
time.sleep(1.5)
deposito = nav.find_element(By.XPATH,'//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[8]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
deposito.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
deposito.send_keys('Serra')
time.sleep(2)
deposito.send_keys(Keys.TAB)
time.sleep(1)
deposito.send_keys(Keys.TAB)
time.sleep(1)
# __________________________________________________________________________________

tabela_bd_estoque = tabela.copy()
valores_recursos = tabela_bd_estoque.iloc[1,9]

recursos = nav.find_element(By.XPATH,'/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1.5)
recursos.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
recursos.send_keys(Keys.BACKSPACE)
time.sleep(1.5)
recursos.send_keys(valores_recursos)
time.sleep(1.5)
data_base.send_keys(Keys.ENTER)
time.sleep(5)
saida_iframe(nav)
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button').click()
time.sleep(4)
iframes(nav)
time.sleep(1)
recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 't')
time.sleep(3) 
recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 'x')

while len(nav.find_elements(By.ID, '_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao')) < 1:
        time.sleep(1)
        print('loading...')

time.sleep(1.5)
saida_iframe(nav)
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]').click()
time.sleep(1.5)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'answers_0')))
nav.find_element(By.ID,'answers_0').click()
time.sleep(1.5)
iframes(nav)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="exporter"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'e')
time.sleep(1.5)
try: 
    while  nav.find_element(By.XPATH, '//*[@id="progressMessageBox"]'):
        print('Carregando ...')
except:
    print('Carregou') 
    time.sleep(1)
    
time.sleep(1.5)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="_download_elt"]')))
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()
time.sleep(3)
# __________________________________________________________________________________________________

ultimoArquivo, caminho = ultimo_arquivo()

tabela_saldo_recursos = pd.read_csv(r'C:/Users/TI/Downloads/' + ultimoArquivo, encoding='iso-8859-1', sep=';')


tabela_saldo_recursos = tabela_saldo_recursos.rename(columns={'="2o. Agrupamento"':'2o. Agrupamento','="Saldo"':'Saldo','="Custo#Total"':'Custo#Total'})

tabela_saldo_recursos = tabela_saldo_recursos.replace('=','',regex=True)
tabela_saldo_recursos.replace('"','',regex=True,inplace=True)

tabela_saldo_recursos['Saldo'] = tabela_saldo_recursos['Saldo'].replace(",",".",regex=True)
tabela_saldo_recursos['Saldo'] = tabela_saldo_recursos['Saldo'].replace("",0,regex=True)
tabela_saldo_recursos['Saldo'] = tabela_saldo_recursos['Saldo'].astype(float)

tabela_saldo_recursos['Custo#Total'] = tabela_saldo_recursos['Custo#Total'].replace(",",".",regex=True)
tabela_saldo_recursos['Custo#Total'] = tabela_saldo_recursos['Custo#Total'].replace("",0,regex=True)
tabela_saldo_recursos['Custo#Total'] = tabela_saldo_recursos['Custo#Total'].astype(float)

tabela_saldo_recursos_lista = tabela_saldo_recursos.fillna('').values.tolist()

sh1.values_clear("'SALDO SERRA'!C2:E")


wks3.update('C2:E', tabela_saldo_recursos_lista)

# _________________________________________________________________________________________________

saida_iframe(nav)
time.sleep(2)
nav.find_element(By.ID,'bt_1892603865').click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_saldo = test_list.loc[test_list[0] == 'Saldos de recursos'].reset_index(drop=True)['index'][0]

lista_menu[click_saldo].click()
time.sleep(2)

iframes(nav)

time.sleep(1.5)

deposito = nav.find_element(By.XPATH,'//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[8]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
deposito.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
deposito.send_keys('Central')
time.sleep(3)
deposito.send_keys(Keys.TAB)
time.sleep(1.5)
deposito.send_keys(Keys.TAB)
time.sleep(1.5)
deposito.send_keys(Keys.CONTROL,Keys.SHIFT + 'x')

while len(nav.find_elements(By.ID, '_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao')) < 1:
        print('loading...')
        time.sleep(1)
        
time.sleep(1.5)
saida_iframe(nav)
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/div[4]/div[3]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span[2]').click()
time.sleep(1.5)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'answers_0')))
nav.find_element(By.ID,'answers_0').click()
time.sleep(1.5)
iframes(nav)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="exporter"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'e')
time.sleep(1.5)
try: 
    while  nav.find_element(By.XPATH, '//*[@id="progressMessageBox"]'):
        print('Carregando ...')
except:
    print('Carregou') 
    time.sleep(1)
    
time.sleep(1.5)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="_download_elt"]')))
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()
time.sleep(3)
# __________________________________________________________________________________________________

ultimoArquivo, caminho = ultimo_arquivo()

tabela_saldo_recursos_central = pd.read_csv(r'C:/Users/TI/Downloads/' + ultimoArquivo, encoding='iso-8859-1', sep=';')

tabela_saldo_recursos_central = tabela_saldo_recursos_central.rename(columns={'="2o. Agrupamento"':'2o. Agrupamento','="Saldo"':'Saldo','="Custo#Total"':'Custo#Total'})

tabela_saldo_recursos_central = tabela_saldo_recursos_central.replace('=','',regex=True)
tabela_saldo_recursos_central.replace('"','',regex=True,inplace=True)

tabela_saldo_recursos_central['Saldo'] = tabela_saldo_recursos_central['Saldo'].replace(",",".",regex=True)
tabela_saldo_recursos_central['Saldo'] = tabela_saldo_recursos_central['Saldo'].replace("",0,regex=True)
tabela_saldo_recursos_central['Saldo'] = tabela_saldo_recursos_central['Saldo'].astype(float)

tabela_saldo_recursos_central['Custo#Total'] = tabela_saldo_recursos_central['Custo#Total'].replace(",",".",regex=True)
tabela_saldo_recursos_central['Custo#Total'] = tabela_saldo_recursos_central['Custo#Total'].replace("",0,regex=True)
tabela_saldo_recursos_central['Custo#Total'] = tabela_saldo_recursos_central['Custo#Total'].astype(float)

print(tabela_saldo_recursos_central.dtypes)

tabela_saldo_recursos_central_lista = tabela_saldo_recursos_central.fillna('').values.tolist()


sh1.values_clear("'ALMOX CENTRAL'!C2:E")


wks2.update('C2:E', tabela_saldo_recursos_central_lista)

# _________________________________________________________________________________________________

hoje = datetime.datetime.now()
hoje = hoje.strftime(format='%d/%m')

wks4.update('G1', 'Saldo do dia - ' + hoje)

