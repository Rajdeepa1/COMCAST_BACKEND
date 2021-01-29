import requests
import json
from requests.auth import HTTPBasicAuth
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

def func_topic(l):

    my_headers = {'Authorization': 'Basic cmFqZGNoYWs6MTFQY20wNDk4QGJoYXZhbg==','Content-Type':'application/x-www-form-urlencoded'}
    print(l)

    param='query={"appId": "91f7727226", "query":"%s","queryLanguage":"SIMPLE","hits": "100","filterQuery": ["table:c3"], "sort": ["date:desc"], "securityRealmId": ["cisco"], "securityPrincipalId": ["rajdchak"], "securityPrincipalName": ["rajdchak"],"fields": ["identifier,company,hwproducttype,swproducttype,swversion,problemdescription,defectcount,linkeddefects,casestate"], "useragent": ["mozilla"], "userIpaddress": ["127.0.0.1"], "clientHost": ["clientHost"], "clientIpaddress": ["127.0.0.1"], "appName": ["COMCAST_Topic_Search"], "appOwner": ["rajdchak"]}&src=TOPIC' % l

    result=list()
    result.append(["SR Number","Headline","Defect","Fixed release","Comments"])
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    






    res_comcast=requests.request("POST","https://wsgi.cisco.com/cso/search/SearchExtService",headers=my_headers,data=param)
    #print(res_comcast.json()["documents"])
    r=2
    for doc in res_comcast.json()["documents"]:
        l=[]
        #print(doc["fields"]["linkeddefects"])
        
        #if "problemdescription" in doc["fields"]:
        l.append(str(doc["fields"]["identifier"])[2:-2])
        
        l.append(str(doc["fields"]["title"])[1:-1])
        
        if(doc["fields"]["defectcount"]!=[0]):
            
            defect=str(doc["fields"]["linkeddefects"])[1:-1]
            l.append(str(doc["fields"]["linkeddefects"])[1:-1])
            
        else:
            
            l.append("N/A")
        if(str(doc["fields"]["text"])[1:-1].find("fixed")!=-1):
            
            l.append(str(doc["fields"]["text"])[str(doc["fields"]["text"])[1:-1].find("fixed"):-1])
        else:
            
            l.append("")
        
        l.append("---")
        
        result.append(l)
    #print(token)
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

