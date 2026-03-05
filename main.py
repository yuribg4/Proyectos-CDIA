from selenium.webdriver import chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
import time
import pandas as pd
from selenium.webdriver.common.by import By

USER="standard_user"
PASSWORD="secret_sauce"

def main():
    service = Service(ChromeDriverManager().install())
    option = webdriver.ChromeOptions()
    option.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(service=service, options=option)

    driver.get("https://www.saucedemo.com/")

    # login
    driver.find_element(By.ID, "user-name").send_keys(USER)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    time.sleep(2)
    driver.find_element(By.ID,"login-button").click()

    # compra
    driver.find_element(By.NAME,"add-to-cart-sauce-labs-bolt-t-shirt").click()
    driver.find_element(By.ID,"add-to-cart-test.allthethings()-t-shirt-(red)").click()

    # carrito
    driver.find_element(By.CLASS_NAME,"shopping_cart_link").click()
    time.sleep(2)

    # obtener datos del carrito
    items = driver.find_elements(By.CLASS_NAME,"cart_item")

    data = []

    for item in items:
        nombre = item.find_element(By.CLASS_NAME,"inventory_item_name").text
        precio = item.find_element(By.CLASS_NAME,"inventory_item_price").text
        cantidad = item.find_element(By.CLASS_NAME,"cart_quantity").text

        data.append({
            "Producto": nombre,
            "Precio": precio,
            "Cantidad": cantidad
        })

    # crear dataframe
    df = pd.DataFrame(data)

    print("\nDataFrame de la compra:\n")
    print(df)

    # checkout
    driver.find_element(By.CLASS_NAME, "checkout_button").click()

    driver.find_element(By.ID, "first-name").send_keys("test")
    driver.find_element(By.ID, "last-name").send_keys("test")
    driver.find_element(By.ID, "postal-code").send_keys("12345")

    driver.find_element(By.ID,"continue").click()
    driver.find_element(By.ID,"finish").click()

    time.sleep(5)
    driver.quit()

if __name__== "__main__":
    main()