# url = "https://www.twse.com.tw/rwd/zh/afterTrading/FMSRFK?date=20230712&stockNo=2330&response=json&_=1689150175809"
# url 2 = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20230713&stockNo=2330&response=json&_=1689213462127"
# import requests

# url = 'https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY'
# rq = requests.get(url,params={
#     "response":"json",
#     "stockNo":"2330",
#     "date":"20230701"

# })

# print(rq.text)

import xml.etree.ElementTree as ET
from openpyxl import Workbook
import requests
import time

def xml_to_dict(root):
    params_dict = {}
    for child in root:
        params_dict[child.tag] = child.text
    return params_dict

def fillsheet(sheet,data,row):
    for column, value in enumerate(data,1):
        sheet.cell(row = row,column = column,value = value)
def returnStrDaylist(startYear,startMonth,endYear,endMonth,day = "01"):
    result = []
    if startYear == endYear: 
        for month in range()(startMonth,endMonth+1) : 
                    month = str(month)
                    if len(month) == 1:
                        month = "0" + month 
                    result.append(str(startYear)+month+day)
        return result
    for year in range(startYear,endYear+1):
        if year == startYear:
                for month in range(startMonth,13) : 
                    month = str(month)
                    if len(month) == 1:
                        month = "0" + month 
                    result.append(str(year)+month+day)
        elif year == endYear :
            for month in range(1,endMonth+1) : 
                    month = str(month)
                    if len(month) == 1:
                        month = "0" + month 
                    result.append(str(year)+month+day)
        else : 
             for month in range(1,13) : 
                    month = str(month)
                    if len(month) == 1:
                        month = "0" + month 
                    result.append(str(year)+month+day)
    return result
def print_params(params_dict):
    for key, value in params_dict.items():
        print(f"{key}: {value}")


xml_file = "data1.xml"
tree = ET.parse("data1.xml")
root = tree.getroot()
data_dict = xml_to_dict(root)
fields= ["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"]
wb = Workbook()
sheet = wb.active
sheet.title = "fields"

fillsheet(sheet,fields,1)
startYear,startmonth = int(data_dict["startYear"]),int(data_dict["startMonth"])
endYear,endMonth = int(data_dict["endYear"]) ,int(data_dict["endmonth"])
yearList = returnStrDaylist(startYear, startmonth, endYear, endMonth)

row =2
for YearMonth in yearList:
    rq = requests.get(data_dict["url"],params={
        "response":"json",
        "date" : YearMonth,
        "stockNo":data_dict["stockNo"]  
    })
    jsonData = rq.json()
    dailyPriceList = jsonData.get("data",[])
    for daliyPrice in dailyPriceList:
        fillsheet(sheet,daliyPrice,row)
        row += 1
     
    time.sleep(3) 

name = data_dict["excelName"]
wb.save(name+".xlsx")

