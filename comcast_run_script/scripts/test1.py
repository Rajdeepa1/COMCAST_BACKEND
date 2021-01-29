from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import os
from bs4 import BeautifulSoup
import time
import getpass
import openpyxl as p
from openpyxl.styles import PatternFill
import subprocess
import sys
import pandas
from io import BytesIO
import xlsxwriter
import socket

def func(l):
    
    
    print("[INFO] Shutting down all chrome processes to use this program.")
    
    #options = webdriver.ChromeOptions()
    #Find your google cookie directory. I have automated this part, if it doesn't work please refer to my video to find your cookie folder.
    #path_to_chrome_cookie="user-data-dir=C:\\Users\\rajdchak\\AppData\\Local\\Google\\Chrome\\User Data"
    #Path to your chrome profile
    #options.add_argument(path_to_chrome_cookie) 
    #chrome_path="/usr/local/bin/chromedriver"
    #driver = webdriver.Chrome(chrome_path, options=options)




    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    #chrome_options.add_argument('--disable_infobars')
    
    driver = webdriver.Chrome(executable_path="chromedriver", options=chrome_options)


    ret_str=""
    result=list()
    result.append(["BugId","Customer Visibility"])
    output = BytesIO()
    for r in l:
        bug=r
        print(bug)
        url="https://cdetsng.cisco.com/summary/#/defect/%s" %(bug)
        driver.get(url)
        time.sleep(2)
        content = driver.page_source
        soup = BeautifulSoup(content,"lxml")
            #print(soup)
        flag=0
        for b in soup.findAll('td'):
            if(flag==1):
                ret_str=ret_str+"---"+str(r)+" "+bug+" "+b.text
                l=[bug,b.text]
                result.append(l)
                flag=0
                break
            if(b.text.find("Is-customer-visible")!=-1):
                flag=1
        #data1 = [['tom', 10], ['nick', 15], ['juli', 14]] 
        #df = pandas.DataFrame(data1, columns=['col1', 'col2'])

        #Convert the data frame to Excel and store it in BytesIO object `workbook`:
        
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # Get some data to write to the spreadsheet.
        data = result

        # Write some test data.
        for row_num, columns in enumerate(data):
            for col_num, cell_data in enumerate(columns):
                worksheet.write(row_num, col_num, cell_data)

        # Close the workbook before sending the data.
        workbook.close()

        # Rewind the buffer.
        output.seek(0)
        
    #return ret_str
    return output

#print(func(['CSCvf69272','CSCvg54149','CSCvj78551','CSCvk00895','CSCvm07353']))