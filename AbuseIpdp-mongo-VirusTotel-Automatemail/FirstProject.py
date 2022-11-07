import requests
import json
import pandas 
import csv
import os
import win32com.client as win32
import time
import pymongo

file_path = str(input('please Enter The File Path: '))
IP_CSV = pandas.read_csv((file_path))

ip=IP_CSV['IP'].tolist()


API_KEY = 'd449e8b1d723cad2d3a2c82455e7346d66804ab1faafbdd33963dc98275926eabe2804191095b440'
url = 'https://api.abuseipdb.com/api/v2/check'

csv_columns = ['ipAddress','isPublic','ipVersion','isWhitelisted','abuseConfidenceScore','countryCode','usageType','isp','domain','hostnames','totalReports','numDistinctUsers','lastReportedAt']

headers = {
    'Accept': 'application/json',
    'Key': API_KEY
}
with open("AbuseIP_results.csv","a", newline='') as filecsv:
    writer = csv.DictWriter(filecsv, fieldnames=csv_columns)
    writer.writeheader()
for i in ip:
    parameters = {
        'ipAddress': i,
        'maxAgeInDays': '90'}

    respnse= requests.get( url=url,headers=headers,params=parameters,verify=False)
    json_Data = json.loads(respnse.content)
    json_main = json_Data["data"]
    with open("AbuseIP_results.csv","a", newline='')as filecsv:
        writer= csv.DictWriter(filecsv,fieldnames=csv_columns)
        writer.writerow(json_main)
client=pymongo.MongoClient("mongodb://localhost:27017")
df=pandas.read_csv("AbuseIP_results.csv")
data=df.to_dict(orient="record")
db=client["AbuseIpdp"]
db.AbuseVT.insert_many(data)

Urls = IP_CSV['IP'].tolist()

API_key = 'YourAPIkey'
url = 'https://www.virustotal.com/vtapi/v2/url/report'

API_key = '417d0360f16cfb585dc7e93f25506fd10498e475409bca93d81a3892e02f76a0'
url = 'https://www.virustotal.com/vtapi/v2/url/report'


parameters = {'apikey': API_key, 'resource': Urls}

for i in Urls:
    parameters = {'apikey': API_key, 'resource': i}

    response= requests.get(url=url, params=parameters,verify=False)
    json_response= json.loads(response.text)
    
    if json_response['response_code'] <= 0:
        with open('not Found result.txt', 'a')  as notfound:
            notfound.write(i) and notfound.write("\tNOT found please Scan it manually\n")
    elif json_response['response_code'] >= 1:

        if json_response['positives'] <= 0:
            with open('Virustotal Clean result.txt', 'a')  as clean:
                clean.write(i) and clean.write("\t NOT malicious \n")
        else:
            with open('Virustotal Malicious result.txt', 'a')  as malicious:
                malicious.write(i) and malicious.write("\t Malicious") and malicious.write("\t this Domains Detectd by   "+ str(json_response['positives']) + "  Solutions\n")

    time.sleep(15)

olApp=win32.Dispatch('Outlook.Application')
olNs=olApp.GetNameSpace('MAPI')

mailItem=olApp.CreateItem(0)
mailItem.Subject='AbuseVT Result'
mailItem.BodyFormat=1
mailItem.Body="Malicious Alert"
mailItem.To='alex.prabhu@optiv.com'

mailItem.Attachments.Add(os.path.join(os.getcwd(),'AbuseIP_results.csv'))
mailItem.Attachments.Add(os.path.join(os.getcwd(),'Virustotal Malicious result.txt'))
mailItem.Attachments.Add(os.path.join(os.getcwd(),'Virustotal Clean result.txt'))
mailItem.Attachments.Add(os.path.join(os.getcwd(),'not Found result.txt'))

mailItem.Display()

mailItem.Save()
mailItem.Send()
print("operation succesful")