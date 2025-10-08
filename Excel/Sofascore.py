from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

# Configuración del navegador (modo headless)
options = Options()
options.headless = True
driver = webdriver.Chrome(options=options)

# Abrir la página de Sofascore
url = "https://www.sofascore.com/es/torneo/futbol/spain/laliga-2/54#id:77558,tab:statistics"
driver.get(url)

# Esperar unos segundos a que cargue todo
time.sleep(5)

# Extraer los datos de la tabla (por ejemplo estadísticas)
table_elements = driver.find_elements(By.CSS_SELECTOR, "div.sc-1v1z6ra-0")  # Selector aproximado
data = []

for table in table_elements:
    rows = table.find_elements(By.CSS_SELECTOR, "div.sc-1y4p9wh-0")  # Cada fila
    for row in rows:
        cols = row.text.split("\n")
        data.append(cols)

# Convertir a DataFrame de Pandas
df = pd.DataFrame(data)
print(df.head())

driver.quit()