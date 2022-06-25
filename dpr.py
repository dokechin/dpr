import openpyxl
from openpyxl import load_workbook
import re
import glob
from tika import parser
import os

import csv

filename="fy-stock-dividend.csv"

dividend = {}
with open(filename,newline="") as csvf:
    data=csv.reader(csvf)
    for i, raw in enumerate(data):
        if i <= 1 : 
            continue
        dividend[raw[0]] = raw[5]

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'].value = '銘柄コード'
sheet['B1'].value = '銘柄名'
sheet['C1'].value = '目標配当性向'
sheet['D1'].value = '配当性向'
sheet['E1'].value = '配当性向差分'
sheet['F1'].value = '抽出結果'

row = 1
files = glob.glob("./pdf/*.pdf")
for file in files:
    print (file)
    filename = os.path.splitext(os.path.basename(file))[0]
    rows = filename.split("_")
    file_data = parser.from_file(file)
    text = file_data["content"]
    if text is None :
        continue
    matchObj = re.search(r'配当性向.*?(\d+\.?\d*)[\%|％]', text)
    if matchObj:
        div = 0
        div_now = 0;
        print(matchObj.group())
        print(matchObj.group(1))
        if matchObj.group(1) is not None :
            try:
                div = float(matchObj.group(1))  # 文字列を実際にfloat関数で変換してみる
            except ValueError:
                div = 0
        if rows[0] in dividend and dividend[rows[0]] is not None :
            try:
                div_now = float(dividend[rows[0]])  # 文字列を実際にfloat関数で変換してみる
            except ValueError:
                div_now = 0
        
        sheet[f"A{row}"] = rows[0]
        sheet[f"B{row}"] = rows[1]
        sheet[f"C{row}"] = div
        sheet[f"D{row}"] = div_now
        sheet[f"E{row}"] =  '' if div_now == 0 else div / div_now
        sheet[f"F{row}"] = matchObj.group()

        row+=1
workbook.save("result.xlsx")


