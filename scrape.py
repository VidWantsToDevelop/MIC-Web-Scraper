import re
import requests
import openpyxl
import urllib.request
from bs4 import BeautifulSoup

global names
names = []
global urls
urls = []
global prices
prices = []
global images
images = []
global allProds
allProds = []


def connect(page):
    soup = None
    products = None

    url = f"https://www.made-in-china.com/multi-search/{srch}/F1/{page}.html"
    print(url)

    attempts = 0

    while attempts < 15:
        attempts += 1
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            webpage = urllib.request.urlopen(req)

            soup = BeautifulSoup(webpage.read(), "html.parser")

            print(soup)

            products = soup.find_all("h2", {"class": "product-name"})

            for product in products:
                allProds.append(product)
                urls.append(product.find_next("a")["href"])
                name = "".join(product.find_next("a").get_text(strip=True))
                name = re.sub(r"((?![A-Z])\w)([A-Z])", r"\1 \2", name)
                names.append(name)

            return soup, products
        except Exception as e:
            print(e)
            print("Error occurred while processing")
            print("Connecting again...")
    return "Timeout"


print("Welcome!")
print("Let's get searching")
inp = input("What are we searching for?: ")

print("And lastly")
pages = input("How many pages to scrape?( be careful here :) ): ")

spInp = re.split(r"\s+", inp)

srch = f"word={spInp[0]}"
for i in range(1, len(spInp)):
    srch += f"%2B{spInp[i]}"


for page in range(int(pages)):
    connect(page)

fname = "result.xlsx"

wbook = openpyxl.Workbook()
sheet = wbook.active


#Managing Excel columns and headers
fhead = openpyxl.styles.Font(
    size = 12,
    bold = True,
)

sheet["B1"] = "Name"
sheet["C1"] = "Price"
sheet["D1"] = "Image"
sheet["E1"] = "URL"

sheet["B1"].font = fhead
sheet["C1"].font = fhead
sheet["D1"].font = fhead
sheet["E1"].font = fhead

print(len(allProds))
print(len(names))
print(len(prices))
print(len(images))
print(len(urls))

print(names)

for i in range(len(allProds)):
    sheet[f"A{i+2}"] = i + 1
    sheet[f"B{i+2}"] = names[i] if i < len(allProds) and len(names) != 0 else "NULL"
    sheet[f"C{i+2}"] = prices[i] if i <= len(allProds) and len(prices) != 0 else "NULL"
    sheet[f"D{i+2}"] = images[i] if i <= len(allProds) and len(images) != 0 else "NULL"

    #URL + stylings
    sheet[f"E{i+2}"].value = "URL"
    sheet[f"E{i+2}"].style = "Hyperlink"
    sheet[f"E{i+2}"].hyperlink = urls[i] if len(urls) != 0 else "NULL"

wbook.save(filename=fname)
