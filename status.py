# %%
import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options


# %%
#Instalando o driver
servico = Service(ChromeDriverManager().install())

#Abrindo o navegador
navegador = webdriver.Chrome(service=servico)

# %%
listas = ["CCRO3","ECOR3","GGPS3","HBSA3","PORT3","RAIL3"]
url = f"https://statusinvest.com.br/acoes/egie3"
navegador.get(url)
navegador.maximize_window()
time.sleep(15)
navegador.find_element(By.CLASS_NAME,'btn-close').click()


# %%
#valuation indicators into a dictionary with dataframes
df_val = {}
for lista in listas:
    url = f"https://statusinvest.com.br/acoes/{lista.lower()}"
    navegador.get(url)
    time.sleep(3)
    navegador.find_element(By.XPATH,'//*[@id="indicators-section"]/div[1]/div[2]/button[2]').click()
    time.sleep(3)
    table = navegador.find_element(By.XPATH,'//*[@id="indicators-section"]/div[3]/div[2]/div/div[2]/div')
    soup = BeautifulSoup(table.get_attribute('outerHTML'), "html.parser")
    
    table_headers = []
    for th in soup.find_all('div','th'):
        table_headers.append(th.text)
        if th.text < "2019":
            break

    table_data = []
    for row in soup.find_all('div','tr'):
        columns = row.find_all('div','td')
        output_row = []
        count = 0
        for column in columns:
            output_row.append(column.text)
            if count > len(table_headers)-2:
                break
            count=count+1
        table_data.append(output_row)
    
    df_val[lista] = pd.DataFrame(table_data, columns=table_headers)
    df_val[lista] = df_val[lista].drop(0)
    
    time.sleep(3)
    name = navegador.find_element(By.XPATH,'//*[@id="indicators-section"]/div[3]/div[2]/div/div[1]')
    name_s = BeautifulSoup(name.get_attribute('outerHTML'), "html.parser")
    name_row = []
    for th in name_s.find_all('h3'):
        name_row.append(th.text)
    df_val[lista].index = name_row


# %%
with pd.ExcelWriter('val_ind.xlsx') as excel_writer:
    for emp in listas:
        df_val[emp].to_excel(excel_writer, sheet_name = emp, index=True)

# %%
navegador.quit()


