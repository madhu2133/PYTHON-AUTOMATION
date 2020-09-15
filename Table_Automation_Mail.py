import pymongo
import requests
import json
import xlrd
from collections import MutableMapping
import smtplib
import email.message
server = smtplib.SMTP('smtp.gmail.com:587')




def convert_flatten(d, parent_key ='', sep ='_'): 
    items = [] 
    for k, v in d.items(): 
        new_key = parent_key + sep + k if parent_key else k 
  
        if isinstance(v, MutableMapping): 
            items.extend(convert_flatten(v, new_key, sep = sep).items()) 
        else: 
            items.append((new_key, v)) 
    return dict(items) 



wb = xlrd.open_workbook("C://....filepath.xlsx")        #Include FilePath of EXCEL DATA
ws = wb.sheet_by_index(0)
mylist = ws.col_values(1)
myset=set(mylist[0:10])



headers = {
    'x-apikey': '17831839697628732369367942'            #Personalized APIKEY
}

url_test ="https://www.virustotal.com/api/v3/ip_addresses/{}"
for i in myset:
    url = url_test.format(i)
    response = requests.get(url, headers=headers)
    decodedResponse = json.loads(response.text)

#Creating a pymongo client
    client = pymongo.MongoClient("mongodb://localhost:27017/")
#Getting the database instance
    db = client['IPCHECK']
#Creating a collection
    coll = db['report']
    #print("dataabase connected")
    coll.insert(decodedResponse,check_keys=False)
 
temp_list = []



x=coll.find({},{'_id':0,'data.id':1,'data.attributes.as_owner':1,'data.attributes.network': 1,'data.attributes.country': 1,'data.attributes.continent': 1}) 
for y in x:
    #print(y)
    changed=convert_flatten(y)
    #print(changed)
    #print(type(changed))

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
      <h1>REPORT OF IP's</h1>
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
msg['Subject'] = 'TABULATION'
 
 
msg['From'] = 'xyz@gmail.com'                  #FROM email address
msg['To'] = 'zbc@gmail.com'                    #TO email address
password = "1234567890"                        #Login Password
msg.add_header('Content-Type', 'text/html')
msg.set_payload(contents)
 
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
 
# Login Credentials for sending the mail
s.login(msg['From'], password)
 
s.sendmail(msg['From'], [msg['To']], msg.as_string())
    