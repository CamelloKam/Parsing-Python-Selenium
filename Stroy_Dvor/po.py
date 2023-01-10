import time
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
import xlsxwriter
list_card_url=[]

#эксель запись

book = xlsxwriter.Workbook(r"C:\Users\Марат\PycharmProjects\Stroy_Dvor\data.xlsx")
page = book.add_worksheet("товары")

row = 0
column = 0

page.set_column("A:A", 50)
page.set_column("B:B", 20)
page.set_column("C:C", 20)
page.set_column("D:D", 20)
#page.set_column("E:E", 20)

#конец эксель записи

#Раскоментировать от начала до конца, чтобы обновить инормацию
## начало
#
# SCROLL_PAUSE_TIME = 3
#
# se = Service("C:/Users/Марат/PycharmProjects/Beskonecno/geckodriver.exe")
# driver = webdriver.Firefox(service= se)
# base_url="https://www.sdvor.com/moscow/category/udarno-rychazhnyj-instrument-8352"
# driver.get(base_url)
#
# # Get scroll height
# last_height =driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
# while True:
#     # Scroll down to bottom
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#
#     # Wait to load page
#     time.sleep(SCROLL_PAUSE_TIME)
#
#     # Calculate new scroll height and compare with last scroll height
#     new_height = driver.execute_script("return document.body.scrollHeight")
#     if new_height == last_height:
#         break
#     last_height = new_height
#
# pageSource = driver.page_source
# fileToWrite = open("page_source.html", "w", encoding="utf-8")
# fileToWrite.write(pageSource)
# fileToWrite.close()
# fileToRead = open("page_source.html", "r", encoding="utf-8")
# fileToRead.close()
# driver.quit()

## конец

#открываю сохраненный html
with open("C:/Users/Марат/PycharmProjects/Stroy_Dvor/page_source.html", encoding="utf8") as file:
    responce=file.read()

#преобразую в lxml
soup = BeautifulSoup(responce, "lxml")
data = soup.findAll("sd-product-grid-item",class_="product-grid-item")

for i in data:
    sylka = i.find("a", class_="product-name").get("href")
    list_card_url.append(sylka)


for card_url in list_card_url:
    responce = requests.get(card_url)
    soup = BeautifulSoup(responce.text, "lxml")
    data = soup.find("cx-page-layout", class_="ProductDetailsPageTemplate")

    #имя товара
    name = data.find("h1").text
    page.write(row, column, name)

    #Цена товара
    coin = data.find("div",class_="price").text
    page.write(row, column + 1, coin)

    # ссылка на фото
    img = data.find("img").get("src")
    page.write(row, column + 2, img)

    #код товара
    kod_tovara = data.find("span",class_="code").text
    page.write(row, column + 3, kod_tovara)

    #наличие товара
    nalcka = data.find("span", class_="total-stock")
    if nalcka is None:
        page.write(row, column + 4, "нет в наличии")
    else:
        page.write(row, column + 4, nalcka.text )


    # #характеристики
    # harakteristik = data.findAll("li",class_="feature")
    # for ii in harakteristik:
    #     tovar = ii.find("li",class_="value").text
    row += 1


book.close()










