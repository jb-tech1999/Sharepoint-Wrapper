import requests
import json


# sharepoint rest wrapper
class Sharepoint:
    # constructor
    def __init__(self, client_id, client_secret, tenant_id, resource, grant_type, tenant, site):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.resource = resource
        self.grant_type = grant_type
        self.tenant = tenant
        self.site = site

    # get token
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
    # get lists

    def setHeaders(self):
        self.headers = {
            'Authorization': f"Bearer {self.token}",
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }

    def getLists(self, guid, filters):

        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists(guid'{guid}')/Items{filters}"
        l = requests.get(list_url, headers=self.headers)
        return l.json()['d']['results']

    def GetEntityname(self, guid):
        entityFullname = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists(guid'{guid}')/listItemEntityTypeFullName"
        entityResponse = requests.get(entityFullname, headers=self.headers)
        return entityResponse.json()['d']['ListItemEntityTypeFullName']

    def CreateItem(self, guid, data):

        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists(guid'{guid}')/Items"
        l = requests.post(list_url, headers=self.headers,
                          data=json.dumps(data))
        return l.json()
