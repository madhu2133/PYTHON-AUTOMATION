import requests
import json
import sys
import xlsxwriter

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class FileManagement:
    def __init__(self):
        self.json_file_name = "FILEPATH.json"        #JSON Filepath
       # self.excel_file_name = "C://Users//Vet//OneDrive//Documents//sample.xlsx"
        self.reset_data()

    def reset_data(self):
        self.array_abuseConfidenceScore = []
        self.array_ipAddress = []
        
    def read_text_file(self):
        try:
            with open(self.json_file_name) as data_file:
                data = json.load(data_file)
                #pprint.pprint(data)
                for each_axis in data["data"]:
                    abuseConfidenceScore = int(each_axis["abuseConfidenceScore"])
                    ipAddress = str(each_axis["ipAddress"])
                    self.array_abuseConfidenceScore.append(abuseConfidenceScore)
                    self.array_ipAddress.append(ipAddress) 
        
        except:
            print("Unexpected error : ", sys.exc_info()[0])
            raise
           

    def save_to_xlsx(self):
        workbook = xlsxwriter.Workbook('sample.xlsx')      #EXCEL FILEPATH TO STORE
        workbook = xlsxwriter.Workbook("sample.xlsx")      #EXCEL FILEPATH TO STORE
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        for index, value in enumerate(self.array_abuseConfidenceScore):
            worksheet.write(index, 0, self.array_abuseConfidenceScore[index])
            worksheet.write(index, 1, self.array_ipAddress[index])
            worksheet.write(index, 0, self.array_abuseConfidenceScore[index]) #column 0
            worksheet.write(index, 1, self.array_ipAddress[index]) #column 1
           

        workbook.close()




# Defining the api-endpoint
url = 'https://api.abuseipdb.com/api/v2/blacklist'

querystring = {
    'confidenceMinimum':'90'
}

headers = {
    'Accept': 'application/json',
    'Key': '123456789876545678765678987654567897654345678876545	',       #PERSONALIZED KEY
}

response = requests.request(method='GET', url=url, headers=headers, params=querystring)

# Formatted output
decodedResponse = json.loads(response.text)
#print(json.dumps(decodedResponse, sort_keys=True, indent=4))
#print(decodedResponse)
s=json.dumps(decodedResponse, sort_keys=True, indent=4)
#print(type(s))
with open("SELECT FILE PATH","w") as f:
    f.write(s)

if __name__ == '__main__':
    file_management = FileManagement()
    file_management.read_text_file()
    file_management.save_to_xlsx()    

EMAIL_ADDRESS = "ABC@gmail.com"                 #ENTER FROM ADDRESS 
PASSWORD = "12345"                              #ENTER PASSWORD
TO_ADDRESS ="XYZ1@gmail.com"                    #ENTER TO ADDRESS


subject = ' AUTOMATION PROJECT '

msg = MIMEMultipart()
msg['From'] =EMAIL_ADDRESS
msg['To'] = TO_ADDRESS
msg['Subject'] = subject

body = "Hi there, I have enclosed the Excel"
msg.attach(MIMEText(body,'plain'))


filename='EXCEL FILE.xlsx'                      #PATH OF EXCEL FILE
attachment  =open(filename,'rb')
part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename= "+ "FILENAME.xlsx")
msg.attach(part)




text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(EMAIL_ADDRESS,PASSWORD)


server.sendmail(EMAIL_ADDRESS,TO_ADDRESS,text)
print("MAIL SUCCESSSFULLY SENT")
server.quit()