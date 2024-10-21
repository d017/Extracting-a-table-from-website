from openpyxl import Workbook
from bs4 import BeautifulSoup

file_name = "files/data.html"
table_name = "files/messier.xlsx"

def get_html(name):
    with open(name, "r", encoding="utf-8") as file:
        data = file.read()
    return data

def load_data(htmlf):
    soup = BeautifulSoup(htmlf, "lxml")

    curr = soup.find("div", id="container")
    curr = curr.find("div", id="mainContent")
    curr = curr.find("div", id="mainContent")
    table = curr.find("table", class_="datatab")

    odds = table.find_all("tr", class_="odd")
    evens = table.find_all("tr", class_="even")

    res = []
    for obj1, obj2 in zip(odds, evens):
        info = obj1.find_all("td")
        appl = [info[0].text.strip(), info[2].text, info[6].text,
                info[7].text, info[8].text, info[10].text.strip()]
        res.append(appl)

        info = obj2.find_all("td")
        appl = [info[0].text.strip(), info[2].text, info[6].text,
                info[7].text, info[8].text, info[10].text.strip()]
        res.append(appl)
    return res

def save_to_workbook(info: list[list], wbname):
    wb = Workbook()
    ws = wb.active
    ws.title = "Messier catalog"
    ws.append(["Number", "Type", "RA", "Decl.", "Const.", "Name"])
    for obj in info:
        ws.append(obj)
    wb.save(wbname)

html_file = get_html(file_name)
information = load_data(html_file)
save_to_workbook(information, table_name)
