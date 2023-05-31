import requests
import json


#sharepoint rest wrapper
class Sharepoint:
    #constructor
    def __init__(self, client_id, client_secret, tenant_id, resource, grant_type, tenant):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.resource = resource
        self.grant_type = grant_type
        self.tenant = tenant

    #get token
    def getToken(self):
        data = {
            'grant_type': self.grant_type,
            'resource': self.resource + "/" + self.tenant + ".sharepoint.com@" + self.tenant_id,
            'client_id': self.client_id + '@' + self.tenant_id,
            'client_secret': self.client_secret,
        }
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        url = f"https://accounts.accesscontrol.windows.net/{self.tenant_id}/tokens/OAuth/2"
        r = requests.post(url, data=data, headers=headers)
        json_data = json.loads(r.text)
        token = json_data['access_token']
        self.token = token
    #get lists
    def getLists(self,site ,guid, filters):
        headers = {
            'Authorization': f"Bearer {self.token}",
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }

        list_url = f"https://{self.tenant}.sharepoint.com/sites/{site}/_api/web/lists(guid'{guid}')/Items{filters}"
        l = requests.get(list_url, headers=headers)
        return l.text
    

