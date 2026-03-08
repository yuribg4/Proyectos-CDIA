from os import name
from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver

from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def main():
    service = Service(ChromeDriverManager().install())
    option = webdriver.ChromeOptions()
    option.add_argument("--window-size=1920,1080")
    driver = Chrome(service=service, options=option)
    driver.get("https://www.mercadolibre.com.mx/ofertas/novedades-de-temporada")
    time.sleep(4) 

    products = driver.find_elements(By.CSS_SELECTOR, ".andes-card.poly-card.poly-card--grid-card")
    product_data = []

    for product in products:
        try:
            name = product.find_element(By.CLASS_NAME, "poly-component__title-wrapper").text
            price = product.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
            product_data.append([name, price])
        except Exception:
            continue

    import pandas as pd
    df = pd.DataFrame(product_data, columns=["Nombre", "Precio"])
    print(df)

    df.to_csv("productos.csv", index=False)
    



if __name__ == '__main__':
    main()