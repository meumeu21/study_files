import time

from selenium  import webdriver
from selenium import common
from selenium.webdriver.common.by import By

import undetected_chromedriver as uc

import openpyxl


def parse(flashcard):
    base = []
    url_flashcard = flashcard.find_element(by=By.CLASS_NAME, value="tile-hover-target.j6y.y6j").get_attribute('href')
    base.append(url_flashcard)
    flashcard_name = flashcard.find_element(By.CLASS_NAME, value="e3m.m3e.e4m.e6m.tsBodyL.j6y.y6j").find_element(by=By.TAG_NAME, value="span").text
    base.append(flashcard_name)
    flashcard_elements = flashcard.find_element(By.CLASS_NAME, value="e3m.m3e.e7m.tsBodyM.j6y").find_element(by=By.TAG_NAME, value="span")
    flashcard_elements = flashcard_elements.text.split('\n')
    names = ['Объем', 'Интерфейсы', 'Скорость чтения, Мб/с', 'Скорость записи, Мб/с']

    isFind = False
    for name in names:
        for flashcard_element in flashcard_elements:
            elements_name = ""
            for char in flashcard_element:
                if char != ':':
                    elements_name += char
                else:
                    break
            if elements_name == name:
                base.append(flashcard_element.split(':')[1])
                isFind = True
                break
        if isFind == False:
            base.append(' ')

    return base


def get_flashcards(driver):
    for _ in range(4):
        scroll(5, 4, driver)
    flashcards = (driver.find_elements(By.CLASS_NAME, value="k4m"))
    return flashcards


def scroll(count, delay, driver):
    for _ in range(count):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(delay)


def main():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument("--disable-extensions")
    options.add_argument('--disable-application-cache')
    options.add_argument('--disable-gpu')
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = uc.Chrome(options=options)
    wb = openpyxl.Workbook()
    ws = wb.active
    header = ['URL', 'Название', 'Объем', 'Интерфейсы', 'Скорость чтения, Мб/с', 'Скорость записи, Мб/с']
    ws.append(header)

    for i in range(1, 4):
        url = 'https://www.ozon.ru/category/usb-fleshki-15755/?category_was_predicted=true&deny_category_prediction=true&page=' + str(i)
        driver.get(url)
        flashcards = get_flashcards(driver)
        for flashcard in flashcards:
            data = parse(flashcard)
            ws.append(data)

    wb.save("flashcards.xlsx")


if __name__ == "__main__":
    main()