from tabula import read_pdf
from openpyxl import Workbook
import wget
import os

workbook = Workbook()
sheet = workbook.active
sheet.column_dimensions["A"].width = 30

name = ["A","B","C","D"]
a = 0
file = "https://www.danielsmetalen.nl/prijslijst_constructiestalen_plaatmaterialen.pdf"
data = read_pdf(wget.download(file), output_format="json")

for i in data[0]["data"]:
        #filename = i['actual']['stationmeasurements']
        a = a + 1
        p = i[0]["text"]
        item = p.split("€ ")
        
        col_A = name[0] + str(a)
        col_B = name[1] + str(a)
        col_C = name[2] + str(a)
        col_D = name[3] + str(a)
        
        if item[-1].replace(",", "").isdigit():
                
                sheet[col_A] = item[0] #data tot excel
                sheet[col_B].number_format = '#,##0.00€' 
                sheet[col_B] = float(item[-1].replace(',', '.'))
                sheet[col_C] = 0
                sheet[col_D].number_format = '#,##0.00€' 
                sheet[col_D] = "=SUM(" + col_B + "*" + col_C + " )"         
                
        else:
                sheet[col_A] = item[0] #data tot excel

# Get total prices and numbers
sheet[name[3] + str(a+1)].number_format = '#,##0.00€' 
sheet[name[3] + str(a+1)] = "=SUM(D4:D" + str(a) + ")"
sheet[name[2] + str(a+1)] = "=SUM(C4:C" + str(a) + ")"
sheet[name[0] + str(a+1)] = "Totaal:"
        
        
x = file.split("/")
os.remove(x[3])

workbook.save(filename="daniels_prijslijst.xls")
