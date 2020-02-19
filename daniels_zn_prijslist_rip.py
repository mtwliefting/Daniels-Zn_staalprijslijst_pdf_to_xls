from tabula import read_pdf
from openpyxl import Workbook
import wget
import os

workbook = Workbook()
sheet = workbook.active
name = ["A","B"]
a = 0


file = "https://www.danielsmetalen.nl/prijslijst_constructiestalen_plaatmaterialen.pdf"

data = read_pdf(wget.download(file), output_format="json")

for i in data[0]["data"]:
        #filename = i['actual']['stationmeasurements']
        a = a + 1
        print()
        p = i[0]["text"]
        item = p.split("€ ")
        
        col_A = name[0] + str(a)
        col_B = name[1] + str(a)
        
        print(col_A)
        
        if item[-1].replace(",", "").isdigit():
                print(item[0] + " € " + item[-1])
                
                sheet[col_A] = item[0] #data tot excel
                sheet[col_B] = item[-1] #data tot excel
        else:
                print(item[0])
                
                sheet[col_A] = item[0] #data tot excel
        
        
        
     
x = file.split("/")
os.remove(x[3])

workbook.save(filename="daniels_prijslijst.xls")
