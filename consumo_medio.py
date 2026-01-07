from selenium import webdriver
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

# UNIDADE 
nome_fazenda = 'santo expedito'


# LEITURA DO EXCEL
df = pd.read_excel('CADASTRO CONSUMOS MEDIO PIMS.xlsx', sheet_name='ATIVOS')

# SELENIUM SETUP
driver = webdriver.Chrome()
actions = ActionChains(driver)
wait = WebDriverWait(driver, 20)

# UNIDADES
fazenda = {"santa lucia":{'cod':'8942926277652785356','nome': 'SANTA LUCIA'},
           "grande leste":{'cod':'5130681048394768998', 'nome': 'GRANDE LESTE' },
           "jatoba":{'cod':'8822813574500442998', 'nome': 'JATOBA' },
           "santo expedito":{'cod':'1363253746237015686', 'nome': 'SANTO EXPEDITO'}}

# LOGIN
driver.get("https://admfranciosi206624.totvsagro.cloudtotvs.com.br/pimsmc/login.jsp")
driver.maximize_window()

wait.until(EC.visibility_of_element_located((By.ID, 'USER'))).send_keys('yan.vieira')
wait.until(EC.visibility_of_element_located((By.ID, 'SENHA'))).send_keys('Y@sf01*')
wait.until(EC.element_to_be_clickable((By.ID, 'Login'))).click()

wait.until(EC.element_to_be_clickable((By.ID, 'nps-button-remove'))).click()

# SELECIONAR UNIDADE
unidade = wait.until(EC.element_to_be_clickable((By.ID, 'UNIADM')))
Select(unidade).select_by_value(fazenda[nome_fazenda.lower()]['cod'])

# NAVEGAÇÃO
time.sleep(1)
tabelas = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[text()='Tabelas']")))
actions.move_to_element(tabelas).perform()

equipamentos = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[normalize-space()='Equipamentos']")))
actions.move_to_element(equipamentos).perform()

driver.find_element(By.LINK_TEXT, "Consumo Médio").click()
wait.until(EC.element_to_be_clickable((By.ID, 'btnAplicar'))).click()

# CADASTRO
sl_erros = [] 
gl_erros = [] 
jb_erros = [] 
se_erros = []


df = df[(df['Fazenda'] == fazenda[nome_fazenda.lower()]['nome']) & (df['Combustivel'] != 'DIESEL S10')].reset_index(drop=True)

for i, row in df.iterrows():
    try:
        wait.until(EC.element_to_be_clickable((By.ID, 'btnIncluir'))).click()

        modelo = driver.find_element(By.ID, 'CODIGO_MODELO')
        modelo.clear()
        modelo.send_keys(str(row['Cod. Modelo']))
        modelo.send_keys(Keys.TAB)

        equipamento = driver.find_element(By.ID, 'CODIGO_EQUIPAMENTO')
        equipamento.clear()
        equipamento.send_keys(str(row['Código']))
        equipamento.send_keys(Keys.TAB)

        material = driver.find_element(By.ID, 'CODIGO_MATERIAL')
        material.clear()
        material.send_keys(str(row['Cod. Com']))
        material.send_keys(Keys.TAB)

        capacidade = driver.find_element(By.ID, 'CAPACIDADE_TANQUE')
        capacidade.clear()
        capacidade.send_keys(str(row['Capacidade']))

        maximo = driver.find_element(By.ID, 'MD_CONS_MAX')
        medio = driver.find_element(By.ID, 'CONSUMO_MEDIO')
        minimo = driver.find_element(By.ID, 'MD_CONS_MIN')
        maximo.clear()
        medio.clear()
        minimo.clear()
        maximo.send_keys(str(row['Maximo']))
        medio.send_keys(str(row['Medio']))
        minimo.send_keys(str(row['Minimo']))

        validade = driver.find_element(By.ID, 'DATA_FINAL_VALIDADE')
        validade.clear()
        validade.send_keys('01/01/2099')

        wait.until(EC.element_to_be_clickable((By.ID, 'btnSalvar'))).click()
        wait.until(EC.element_to_be_clickable((By.ID, 'messageBtnOK'))).click()

        wait.until(EC.element_to_be_clickable((By.ID, 'btnIncluir')))
        time.sleep(0.4)

        print(f"{row['Código']} - Cadastrado com sucesso!")

    except Exception as e:
        print(f"Erro {e} | Equipamento ({row['Código']})")
        if df.loc[i, 'Fazenda'] == 'SANTA LUCIA': sl_erros.append(df.loc[i, 'Código']) 
        if df.loc[i, 'Fazenda'] == 'GRANDE LESTE': gl_erros.append(df.loc[i, 'Código']) 
        if df.loc[i, 'Fazenda'] == 'JATOBA': jb_erros.append(df.loc[i, 'Código']) 
        if df.loc[i, 'Fazenda'] == 'SANTO EXPEDITO': se_erros.append(df.loc[i, 'Código'])
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'messageBtnOK'))).click()
            wait.until(EC.element_to_be_clickable((By.ID, 'btnCancelar'))).click()
        except:
            pass

# EXPORTAR ERROS
if sl_erros:
    pd.DataFrame(sl_erros).to_excel(f'ERROS SANTA LUCIA.xlsx', index=False)
if gl_erros:
    pd.DataFrame(gl_erros).to_excel(f'ERROS GRANDE LESTE.xlsx', index=False) 
if jb_erros:
    pd.DataFrame(jb_erros).to_excel(f'ERROS JATOBA.xlsx', index=False) 
if se_erros:
    pd.DataFrame(se_erros).to_excel(f'ERROS SANTO EXPEDITO.xlsx', index=False)

print("Processo finalizado.")