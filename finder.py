from selenium import webdriver



import xlsxwriter as x

from time import sleep
import datetime
start = datetime.datetime.now()
import threading
from path import PATH
driver = webdriver.Chrome(PATH)

book = x.Workbook('stocks.xlsx')
sheet = book.add_worksheet() 
sheet.col_formats


def therds():
    threading.Thread(target=get).start()
    threading.Thread(target=heads).start()


def get():
    driver.get('https://www.investing.com/equities/')



def heads():
    
    sheet.set_column(
    "A:A",
    15
    )

    sheet.write('A1', "Stock")
    sheet.write("B1", "Cost")
    sheet.write('C1', "High")
    sheet.write("D1", "Low")
    sheet.write("E1", "Price Change")
    sheet.write("F1", "Chage In %")
    sheet.write("G1", "Vol")



def getdata():
    for ii in range(30):
        
        i = ii + 1
        ii = i + 1


        xpath1 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[2]/a")
        xpath2 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[3]")
        xpath3 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[4]")
        xpath4 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[5]")
        xpath5 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[6]")
        xpath6 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[7]")
        xpath7 = (f"/html/body/div[5]/section/div[8]/table/tbody/tr[{i}]/td[8]")
        stock = driver.find_element_by_xpath(xpath1).text
        cost = driver.find_element_by_xpath(xpath2).text
        high = driver.find_element_by_xpath(xpath3).text
        low = driver.find_element_by_xpath(xpath4).text
        change1 = driver.find_element_by_xpath(xpath5).text
        change2 = driver.find_element_by_xpath(xpath6).text
        vol = driver.find_element_by_xpath(xpath7).text

        sheet.write(f'A{ii}', stock)
        sheet.write(f'B{ii}', cost)
        sheet.write(f'C{ii}', high)
        sheet.write(f'D{ii}', low)
        sheet.write(f'E{ii}', change1)
        sheet.write(f'F{ii}', change2)
        sheet.write(f'G{ii}', vol)




therds()
getdata()

book.close()
driver.quit()



finish = datetime.datetime.now()



print(f'Finished In {finish - start}')