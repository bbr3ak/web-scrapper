import time

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait

# options = webdriver.ChromeOptions()
# options.add_experimental_option("excludeSwitches", ["enable-automation"])
# options.add_experimental_option("useAutomationExtension", False)
# options.add_argument("--disable-blink-features=AutomationControlled")

d = webdriver.Chrome()

try:
    d.get("https://hh.ru/search/vacancy?text=&area=1")
    time.sleep(3)

    vlinks = d.find_elements(
        By.XPATH, "//a[contains(@href, '/vacancy/') and @target='_blank']"
    )
    vdata = []

    for i in range(min(5, len(vlinks))):
        vl = vlinks[i]
        vurl = vl.get_attribute("href")

        d.execute_script(f"window.open('{vurl}', '_blank');")
        time.sleep(2)

        d.switch_to.window(d.window_handles[1])
        time.sleep(3)

        try:
            t_el = d.find_element(By.CSS_SELECTOR, 'h1[data-qa="vacancy-title"]')
            title = t_el.text.strip()
        except NoSuchElementException:
            title = "Не указано"

        # Зарплата извлекается через JS, а не через стандартные селекторы Selenium.
        # Причина: верстка hh.ru нестабильна — классы у блока зарплаты генерируются
        # динамически (например, "compensation-row--abc123"), и CSS/XPath-селекторы
        # ломаются при каждом обновлении сайта. Проще один раз написать JS, который
        # ищет элемент по содержимому (символ ₽), чем поддерживать хрупкие селекторы.
        js_get_salary = """
                function getSalary() {
                    let el = document.querySelector('div[class^="compensation-row"]');

                    if (!el) {
                        const elements = Array.from(document.querySelectorAll('span, div, p'));
                        el = elements.find(e =>
                            e.innerText.includes('₽') &&
                            e.innerText.length < 50 &&
                            e.getBoundingClientRect().top < 500
                        );
                    }

                    if (el) {
                        return el.innerText.replace(/\\s+/g, ' ').trim();
                    }
                    return "Зарплата не найдена";
                }
                return getSalary();
                """

        try:
            salary = d.execute_script(js_get_salary)
        except Exception as e:
            salary = f"Ошибка JS: {e}"

        try:
            desc_el = d.find_element(
                By.CSS_SELECTOR, 'div[data-qa="vacancy-hiring-formats"]'
            )
            desc = desc_el.text.strip()
        except NoSuchElementException:
            desc = "Не указано"

        skills = ""
        skills_elements = d.find_elements(
            By.CSS_SELECTOR, "li[data-qa='skills-element']"
        )
        for se in skills_elements:
            skill = se.find_element(By.CSS_SELECTOR, "div > div").text.strip()
            skills += skill + ", "
        if not skills:
            skills = "Не указана"
        else:
            skills = skills[0:-2]

        vdata.append(
            {
                "Название вакансии": title,
                "Зарплата": salary,
                "Краткое описание": desc,
                "Ключевые навыки": skills,
                "Ссылка": vurl,
            }
        )

        d.close()
        d.switch_to.window(d.window_handles[0])
        time.sleep(4)

    if vdata:
        df = pd.DataFrame(vdata)

        with pd.ExcelWriter("pr1hh1.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Вакансии")

            worksheet = writer.sheets["Вакансии"]

except Exception as e:
    print(f"Критическая ошибка: {e}")
    import traceback

    traceback.print_exc()

finally:
    input("\nНажмите Enter для закрытия браузера...")
    d.quit()
