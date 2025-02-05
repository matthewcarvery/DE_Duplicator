import requests
import json
from time import time

def generate_access_token(client_id: str, clientsecret: str, subdomain) -> str:
    auth_base_url = f'https://{subdomain}.auth.marketingcloudapis.com/v2/token'
    headers = {'content-type': 'application/json'}
    payload = {'grant_type': 'client_credentials', 'client_id': client_id, 'client_secret': clientsecret}
    authentication_response = requests.post(url=auth_base_url, data=json.dumps(payload), headers=headers).json()

    if 'access_token' not in authentication_response:
        raise Exception(f'Unable to validate (ClientID/ClientSecret): {repr(authentication_response)}')
    
    access_token = authentication_response['access_token']
    expires_in = time() + authentication_response['expires_in']

    return(access_token, expires_in)