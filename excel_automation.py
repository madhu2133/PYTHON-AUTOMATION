import smtplib
import requests
import json
import sys
import xlsxwriter
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class FileManagement:
    def __init__(self):
        self.json_file_name = "LOCATION PATH//ip_score.json"    #JSON FILE WITH ABUSECONFIDENCESCORE AND IPADDRESS
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
           

    def save_to_xlsx(self):                      #FUNCTION TO STORE JSON DATA IN EXCEL NAMED AS abuse_score
        workbook = xlsxwriter.Workbook('C://Users//Vet//OneDrive//Documents//abuse_score.xlsx')
        workbook = xlsxwriter.Workbook("C://Users//Vet//OneDrive//Documents//abuse_score.xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        for index, value in enumerate(self.array_abuseConfidenceScore):
            worksheet.write(index, 0, self.array_abuseConfidenceScore[index])
            worksheet.write(index, 1, self.array_ipAddress[index])
            worksheet.write(index, 0, self.array_abuseConfidenceScore[index]) #column 0
            worksheet.write(index, 1, self.array_ipAddress[index]) #column 1
           

        workbook.close()

url = 'https://api.abuseipdb.com/api/v2/blacklist'   #DEFINING THE API-ENDPOINT

querystring = {
    'confidenceMinimum':'85'                #QUERY FOR IP’s WITH "abuseConfidenceScore" GREATER THAN 85 FROM AbuseIPDb
}

headers = {
    'Accept': 'application/json',
    'Key': '23dwthbnn984739nwdkdbsmxnx',    #PERSONALIZED PUBLIC API KEY FOR AbuseIPDb ACCOUNT
}

response = requests.request(method='GET', url=url, headers=headers, params=querystring)
decodedResponse = json.loads(response.text)         #FORMATTED OUTPUT
s=json.dumps(decodedResponse, sort_keys=True, indent=4)
with open("PATH_LOCATION//ip_score.json","w") as f:  #OPENING JSON FILE TO STORE ABUSECONFIDENCESCORE AND IPADDRESS
    f.write(s)
    
if __name__ == '__main__':                   #FUNCTION CALLING TO CONVERT JSON DATA FILE INTO EXCEL SHEET
    file_management = FileManagement()
    file_management.read_text_file()
    file_management.save_to_xlsx()

EMAIL_ADDRESS = "FROM@gmail.com"             #FROM ADDRESS
PASSWORD = "FROMPASSWORD"                    #FROM PASSWORD
TO_ADDRESS ="TO@gmail.com"                   #TO PASSWORD
subject = ' AUTOMATION PROJECT '
msg = MIMEMultipart()
msg['From'] =EMAIL_ADDRESS
msg['To'] = TO_ADDRESS
msg['Subject'] = subject
body = "Hi, I have enclosed the Excel containing the AbuseConfidence Score of recent IP's"
msg.attach(MIMEText(body,'plain'))
filename='PATH_LOCATION//abuse_score.xlsx'     #EXCEL SHEET LOACTION IN LOCAL
attachment  =open(filename,'rb')
part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename= "+ "ABUSE_SCORE.xlsx")
msg.attach(part)                               #EXCEL SHEET ATTACHED
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(EMAIL_ADDRESS,PASSWORD)
server.sendmail(EMAIL_ADDRESS,TO_ADDRESS,text) #SENDING AUTOMATED MAIL
print("MAIL SUCCESSSFULLY SENT")
server.quit()