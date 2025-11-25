import pandas as pd
import xmltodict
import json
import requests
import sys
import os
import pyexcel as p
from auth import generate_access_token
import secret
from tkinter import filedialog
sys.setrecursionlimit(4000)

accountname = "MWG_DE"
client_id = secret.account[accountname]['client_id']
subdomain = secret.account[accountname]['subdomain']
MID = secret.account[accountname]['MID']
clientsecret = secret.account[accountname]['clientsecret']
resturl = f'https://{subdomain}.rest.marketingcloudapis.com/'
soapurl = f'https://{subdomain}.soap.marketingcloudapis.com/Service.asmx'
token = generate_access_token(client_id, clientsecret, subdomain)


folderID = '401456'

#MainDETemplate = 'AAA_Data_Spec'

def getDEfields(DEVal, DEProp):
    payload = f"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\" xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\" xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">\n    <s:Header>\n        <a:Action s:mustUnderstand=\"1\">Retrieve</a:Action>\n        <a:To s:mustUnderstand=\"1\">{soapurl}</a:To>\n        <fueloauth xmlns=\"http://exacttarget.com\">{token[0]}</fueloauth>\n    </s:Header>\n     <s:Body xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n        <RetrieveRequestMsg xmlns=\"http://exacttarget.com/wsdl/partnerAPI\">\n         <RetrieveRequest>\n            <ObjectType>DataExtensionField</ObjectType>\n                <Properties>DefaultValue</Properties>\n                <Properties>FieldType</Properties>\n                <Properties>IsPrimaryKey</Properties>\n                <Properties>IsRequired</Properties>\n                <Properties>CustomerKey</Properties>\n                <Properties>MaxLength</Properties>\n                <Properties>Name</Properties>\n                <Properties>ObjectID</Properties>\n                <Properties>Ordinal</Properties>\n                <Properties>StorageType</Properties>\n            <Filter xsi:type=\"SimpleFilterPart\">\n               <Property>{DEProp}</Property>\n               <SimpleOperator>equals</SimpleOperator>\n               <Value>{DEVal}</Value>\n            </Filter>\n            <QueryAllAccounts>true</QueryAllAccounts>\n            <Retrieves />\n            <Options>\n               <SaveOptions />\n               <IncludeObjects>true</IncludeObjects>\n            </Options>\n            </RetrieveRequest>\n        </RetrieveRequestMsg>\n    </s:Body>\n</s:Envelope>"
    headers = {"content-type": "text/xml"}
    response = requests.request("POST", soapurl, headers=headers, data=payload)
    o = xmltodict.parse(response.text)
    if "Results" in o["soap:Envelope"]["soap:Body"]["RetrieveResponseMsg"].keys():
        q = o["soap:Envelope"]["soap:Body"]["RetrieveResponseMsg"]["Results"]
        nq = sorted(q, key=lambda d: int(d['Ordinal']))
        return(nq)
    else:
        print("Source DE not found. Please check with Ashley Kam to resolve")
        sys.exit()

def getDEProps(obID):
    payload = f"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\" xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\" xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">\n    <s:Header>\n        <a:Action s:mustUnderstand=\"1\">Retrieve</a:Action>\n        <a:To s:mustUnderstand=\"1\">{soapurl}</a:To>\n        <fueloauth xmlns=\"http://exacttarget.com\">{token[0]}</fueloauth>\n    </s:Header>\n    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n        <RetrieveRequestMsg xmlns=\"http://exacttarget.com/wsdl/partnerAPI\">\n            <RetrieveRequest>\n                <ObjectType>DataExtension</ObjectType>\n                <Properties>ObjectID</Properties>\n                <Properties>CustomerKey</Properties>\n                <Properties>Name</Properties>\n                <Properties>IsSendable</Properties>\n                <Properties>SendableSubscriberField.Name</Properties>\n                <Properties>SendableSubscriberField.Name</Properties>\n                <Filter xsi:type=\"SimpleFilterPart\">\n                    <Property>ObjectID</Property>\n                    <SimpleOperator>equals</SimpleOperator>\n                    <Value>{obID}</Value>\n                </Filter>\n            </RetrieveRequest>\n        </RetrieveRequestMsg>\n    </s:Body>\n</s:Envelope>"
    headers = {"content-type": "text/xml"}
    response = requests.request("POST", soapurl, headers=headers, data=payload)
    return(xmltodict.parse(response.text))

def makeDE(DEName, newList, folderID, test, SF):
    uppayloadstart = f"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\" xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\" xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">\n    <s:Header>\n        <a:Action s:mustUnderstand=\"1\">Create</a:Action>\n        <a:To s:mustUnderstand=\"1\">{soapurl}</a:To>\n<fueloauth xmlns=\"http://exacttarget.com\">{token[0]}</fueloauth>\n    </s:Header>\n    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n  <CreateRequest xmlns=\"http://exacttarget.com/wsdl/partnerAPI\">\n<Objects xsi:type=\"DataExtension\">\n <Client>\n <ID>{MID}</ID> \n </Client> \n<CategoryID>{folderID}</CategoryID>\n                <CustomerKey></CustomerKey>\n  <Name>{DEName}</Name>\n             <IsTestable>true</IsTestable>\n             <IsSendable>true</IsSendable>\n        <SendableDataExtensionField>\n       <CustomerKey>{SF}</CustomerKey>\n          <Name>{SF}</Name>\n <FieldType>Text</FieldType>\n </SendableDataExtensionField>\n <SendableSubscriberField>\n                    <Name>Subscriber Key</Name>\n   <Value></Value>\n  </SendableSubscriberField>\n   <Fields>\n"
    uploadmid = ""
    for x in  range(len(newList)):
        if test == 1:
            priKey = 'false'
        else:
            priKey = newList[x]['IsPrimaryKey']
        DefaultVal = newList[x]['DefaultValue'] if newList[x]['DefaultValue'] != None else ""
        if ('MaxLength' not in newList[x]) or (newList[x]['FieldType'] == "Text" and newList[x]['MaxLength'] == 0) or (newList[x]['FieldType'] == "Number") or (newList[x]['FieldType'] == "Date") or (newList[x]['FieldType'] == "Boolean") :
            uploadmid += f"<Field>\n   <CustomerKey>{newList[x]['Name']}</CustomerKey>\n   <Name>{newList[x]['Name']}</Name>\n <FieldType>{newList[x]['FieldType']}</FieldType>\n  <IsRequired>{newList[x]['IsRequired']}</IsRequired>\n   <IsPrimaryKey>{priKey}</IsPrimaryKey>\n<DefaultValue>{DefaultVal}</DefaultValue>\n</Field>\n "
        else:
            uploadmid += f"<Field>\n   <CustomerKey>{newList[x]['Name']}</CustomerKey>\n   <Name>{newList[x]['Name']}</Name>\n <FieldType>{newList[x]['FieldType']}</FieldType>\n  <IsRequired>{newList[x]['IsRequired']}</IsRequired>\n   <IsPrimaryKey>{priKey}</IsPrimaryKey>\n<DefaultValue>{DefaultVal}</DefaultValue>\n<MaxLength>{newList[x]['MaxLength']}</MaxLength></Field>\n"
    uppayloadend = "</Fields> \n     </Objects> \n        </CreateRequest> \n   </s:Body> \n </s:Envelope>"
    uppayload = uppayloadstart + uploadmid + uppayloadend
    headers = {"content-type": "text/xml"}
    response = requests.request("POST", soapurl, headers=headers, data=uppayload)
    m = xmltodict.parse(response.text)
    if m["soap:Envelope"]["soap:Body"]["CreateResponse"]["Results"]["StatusCode"] != "OK":
        print(f"Error creating DE. Please ensure a DE with the name '{DEName}' does not already exist.")
        return("error")
    else:
        return(m)

def postdata(items, ukey):
    uaccess_token, uexpire = generate_access_token(client_id, clientsecret, subdomain)
    uheaders = {'authorization': f'Bearer {uaccess_token}', 'content-type': 'application/json'}
    payload = {"items": items}
    urest_url = f'{resturl}data/v1/async/dataextensions/key:{ukey}/rows'
    insert_request = requests.post(url=f'{urest_url}', data=json.dumps(payload), headers=uheaders)
    return(insert_request)

def defineSheets(newList, method):
    folderID = input("Enter folder ID: ")
    fileName= filedialog.askopenfilename()
    projectName = os.path.splitext(os.path.basename(fileName))[0]
    book = p.get_book(file_name=fileName)
    sheets = book.to_dict()
    sheetList = sheets.keys()
    if method == "1":
        SendField = "Email Address"
        m = makeDE(projectName, newList, folderID, 0, SendField)
    if method == "2":
        SendField = input("Enter Field Name that relates to Subscribers on Subscriber Key: ")   
    for x in sheetList:
        o = makeDE(x + "_" + projectName, newList, folderID, 1, SendField)
        if o != "error":
            print(o["soap:Envelope"]["soap:Body"]["CreateResponse"]["Results"]['StatusMessage'])
            oID = o["soap:Envelope"]["soap:Body"]["CreateResponse"]["Results"]['NewObjectID']
            u = getDEProps(oID)
            ukey = u["soap:Envelope"]["soap:Body"]["RetrieveResponseMsg"]["Results"]["CustomerKey"]
            excel_data_df = pd.read_excel(fileName, sheet_name=x)
            excel_data_df = excel_data_df.fillna('')
            #print(excel_data_df)
            json_str = excel_data_df.to_dict(orient='records')
            if len(json_str) > 0:
                print(postdata(json_str, ukey))


print("DE Source:")
print("1 - Master DE from Engage")
print("2 - Custom DE")
method = input("DE Source: ")
if method == "1":
    MainDETemplate = 'CA30ABEB-04C7-4EB6-9CD2-9112D904E058'
elif method == "2":
    MainDETemplate = input("Enter CustomerKey of DE: ")
newList = getDEfields(MainDETemplate, 'DataExtension.CustomerKey')
ds = defineSheets(newList, method)    



#json_str = [{"SubscriberKey": "Key1", "IsActive": 1, "FirstName": "Steve", "LastName": "Smith"}]
#ukey = "PM_API_TEST"
#postdata(json_str, ukey)
