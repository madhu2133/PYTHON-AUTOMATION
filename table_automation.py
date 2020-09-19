import pymongo
import requests
import json
import xlrd
import smtplib
import email.message
from collections import MutableMapping
server = smtplib.SMTP('smtp.gmail.com:587')


def convert_flatten(d, parent_key ='', sep ='_'):                    #NESTED DICT TO FLATTEN DICT
    items = [] 
    for k, v in d.items(): 
        new_key = parent_key + sep + k if parent_key else k 
  
        if isinstance(v, MutableMapping): 
            items.extend(convert_flatten(v, new_key, sep = sep).items()) 
        else: 
            items.append((new_key, v)) 
    return dict(items) 

wb = xlrd.open_workbook("EXCEL_PATH_LOCATION//abuse_score.xlsx")       #EXCEL SHEET LOACTION OF 'ABUSECONFIDENCESCORE' AND 'IP ADDRESS' IN LOCAL
ws = wb.sheet_by_index(0)
mylist = ws.col_values(1)
myset=set(mylist[0:10])

headers = {
    'x-apikey': 'n12xjiwhnknxjkxndi12892',                             #VIRUSTOTAL PUBLIC API KEY
}

url_test ="https://www.virustotal.com/api/v3/ip_addresses/{}"
for i in myset:
    url = url_test.format(i)
    response = requests.get(url, headers=headers)
    decodedResponse = json.loads(response.text)
    client = pymongo.MongoClient("mongodb://localhost:27017/")         #CREATING A PYMONGO CLIENT
    db = client['IPCHECK']                                             #GETTING THE DATABASE INSTANCE NAMED 'IPCHECK'
    coll = db['report']                                                #CREATING A COLLECTION NAMED 'REPORT'
    coll.insert(decodedResponse,check_keys=False)                      
 
temp_list = []
x=coll.find({},{'_id':0,'data.id':1,'data.attributes.as_owner':1,'data.attributes.network': 1,'data.attributes.country': 1,'data.attributes.continent': 1}) 
for y in x:
    changed=convert_flatten(y)
    a = "<tr><td>%s</td>"%changed['data_attributes_as_owner']
    temp_list.append(a)
    b = "<td>"  " %s</td>"%changed['data_attributes_continent']
    temp_list.append(b)
    c = "<td>"  " %s</td>"%changed['data_attributes_country']
    temp_list.append(c)
    d = "<td>"  "%s</td>"%changed['data_attributes_network']
    temp_list.append(d)
    e = "<td>"  " %s</td></tr>"%changed['data_id']
    temp_list.append(e)
    
#CREATING A HTML TO TABULATE THE DETAILS OF IP    
contents = '''<!DOCTYPE html>
<html>
   <head>
      <style>
         table {  
    width: 640px; 
    border-collapse: collapse; 
    border-spacing: 0;
                }
    td, th { border: 1px solid #CCC; }  
      </style>
   </head>

   <body>
      <h1>DETAILS OF IP</h1>
      <table>
         <tr>
            <th>DATA_ATTRIBUTES_AS_OWNER</th>
            <th>DATA_ATTRIBUTES_COUNTRY</th>
            <th>DATA_ATTRIBUTES_CONTINENT</th>
            <th>DATA_ATTRIBUTES_NETWORKS</th>
            <th>IPADDRESS</th>
         </tr>
         <tr>
             %s
         </tr>
      </table>
   </body>
</html>
''' %' '.join(map(str, temp_list))

msg = email.message.Message()
msg['Subject'] = 'VIRUSTOTAL TABULATION REPORT'
msg['From'] = 'FROM@gmail.com'                        #FROM ADDRESS
msg['To'] = 'TO@gmail.com'                            #TO ADDRESS
password = "FROMPASSWORD"                             #FROM PASSWORD  
msg.add_header('Content-Type', 'text/html')
msg.set_payload(contents)                             #TABULATION IN BODY OF THE MAIL
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login(msg['From'], password)                        #LOGIN CREDENTIALS FOR SENDING THE MAIL 
s.sendmail(msg['From'], [msg['To']], msg.as_string()) #EMAIL SENT SUCCESSFULLY
    
