from auth import generate_access_token
import requests
import secret

accountname = "AAA_MWG"
client_id = secret.account[accountname]['client_id']
subdomain = secret.account[accountname]['subdomain']
MID = secret.account[accountname]['MID']
clientsecret = secret.account[accountname]['clientsecret']
resturl = f'https://{subdomain}.rest.marketingcloudapis.com/'
soapurl = f'https://{subdomain}.soap.marketingcloudapis.com/Service.asmx'
uaccess_token, uexpire = generate_access_token(client_id, clientsecret, subdomain)
print(uaccess_token)


url = resturl + '/hub/v1/dataevents/key:D75B1A3B-63AA-44B3-9C18-FAD376D0E3C8/rowset'
headers = {'authorization': f'Bearer {uaccess_token}', 'content-type': 'application/json'}
body = """{
  "items": [
    {
      "MemberId": "9dfc9341-d9d3-44f3-8624-f4c2dc87b4b3",
      "FirstName": "Vijay",
      "Surname": "Lahiri",
      "RewardsPoints": 7742,
      "RewardsTier": 1,
      "Area": "Haringey"
    }
  ]
}"""

req = requests.post(url, headers=headers, data=body)

print(req.status_code)
print(req.headers)
print(req.text)