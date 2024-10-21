from openpyxl import load_workbook, Workbook

filename = "files/messier.xlsx"
table = load_workbook(filename)
sheet = table.active

print(sheet["A2"].value)
sheet["A2"].value = "o"
print(sheet["A2"].value)
