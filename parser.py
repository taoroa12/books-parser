import requests
from bs4 import BeautifulSoup
import openpyxl

BASE_URL = "https://books.toscrape.com/catalogue/"
URL = "https://books.toscrape.com/catalogue/page-{}.html"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Книги"
ws.append(["Название", "Цена", "Рейтинг"])

RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

print("Парсинг начался...")

for page in range(1, 6):
    response = requests.get(URL.format(page))
    soup = BeautifulSoup(response.text, "html.parser")
    books = soup.find_all("article", class_="product_pod")
    
    for book in books:
        title = book.h3.a["title"]
        price = book.find("p", class_="price_color").text.strip()
        rating = RATINGS[book.p["class"][1]]
        ws.append([title, price, rating])
    
    print(f"Страница {page} готова")

wb.save("books.xlsx")
print("Готово! Файл books.xlsx сохранён")