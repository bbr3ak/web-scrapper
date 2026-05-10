import time

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-blink-features=AutomationControlled")

d = webdriver.Chrome(options=options)

try:
    d.get(
        "https://www.wildberries.ru/catalog/0/search.aspx?search=%D1%80%D0%B5%D0%B7%D0%B8%D0%BD%D0%BE%D0%B2%D1%8B%D0%B5%20%D1%83%D1%82%D0%BE%D1%87%D0%BA%D0%B8"
    )
    time.sleep(30)

    wbarts = d.find_elements(By.XPATH, "//article[contains(@class, 'product-card')]")
    wbdata = []

    for wbart in wbarts:
        try:
            t_el = wbart.find_element(By.CLASS_NAME, "product-card__brand")
            title = t_el.text.strip()
        except NoSuchElementException:
            title = "Не указано"

        try:
            img_el = wbart.find_element(
                By.XPATH, "//img[contains(@class, 'j-thumbnail')]"
            )
            img = img_el.get_attribute("src")
        except NoSuchElementException:
            img = "Не указано"

        try:
            desc_el = wbart.find_element(By.CLASS_NAME, "product-card__name")
            desc = desc_el.text.strip()

        except NoSuchElementException:
            desc = "Не указано"

        try:
            rate_el = wbart.find_element(
                By.XPATH, "//span[contains(@class, 'address-rate-mini')]"
            )
            rate = rate_el.text.strip()
        except NoSuchElementException:
            rate = "Не указано"

        try:
            price_el = wbart.find_element(By.CLASS_NAME, "price__lower-price")
            price = price_el.text.strip()

        except NoSuchElementException:
            price = "Не указано"

        wbdata.append(
            {
                "Называние": title,
                "Описание": desc,
                "Рейтинг": rate,
                "Цена": price,
                "Изображение": img,
            }
        )

        if len(wbdata) >= 5:
            break

    print(wbdata)

    if wbdata:
        df = pd.DataFrame(wbdata)

        with pd.ExcelWriter("pr1wb.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Маркетплейс")

            worksheet = writer.sheets["Маркетплейс"]

except Exception as e:
    print(f"Критическая ошибка: {e}")
    import traceback

    traceback.print_exc()

finally:
    input("\nНажмите Enter для закрытия браузера...")
    d.quit()
