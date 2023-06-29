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

    def getLists(self, listTitle, filters):

        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists/GetByTitle('{listTitle}')/Items{filters}"
        l = requests.get(list_url, headers=self.headers)
        return l.json()['d']['results']

    def GetEntityname(self, listTitle):
    #def GetEntityname(self):
        entityFullname = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists/GetByTitle('{listTitle}')/listItemEntityTypeFullName"
        
        entityResponse = requests.get(entityFullname, headers=self.headers)
        if entityResponse.status_code == 404:
            return {"error": "List not found","Message":"Create list first"}
        return entityResponse.json()['d']['ListItemEntityTypeFullName']
    
    def GetListFields(self, listTitle):
        FieldList = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists/GetByTitle('{listTitle}')/fields?$filter=Hidden eq false and ReadOnlyField eq false"
        FieldResponse = requests.get(FieldList, headers=self.headers)
        return FieldResponse.json()['d']['results']



    def CreateItem(self, listTitle, data):

        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists/GetByTitle('{listTitle}')/Items"

        l = requests.post(list_url, headers=self.headers,
                          data=json.dumps(data))
        if l.status_code == 400:
            return {"error": "You messed up", "Message": l.json()}
        
        return l.json()

    def createNewList(self, listName):
        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists"
        data = {
            '__metadata': {'type': 'SP.List'},
            'AllowContentTypes': True,
            'BaseTemplate': 100,
            'ContentTypesEnabled': True,
            'Description': 'My list description',
            'Title': listName
        }

        l = requests.post(list_url, headers=self.headers,
                          data=json.dumps(data))
        return l.json()

    def createNewFields(self, fieldnames, listTitle):
        list_url = f"https://{self.tenant}.sharepoint.com/sites/{self.site}/_api/web/lists/GetByTitle('{listTitle}')/fields"
        for fieldname in fieldnames:
            try:
                data = {
                    '__metadata': {'type': 'SP.Field'},
                    'FieldTypeKind': 2,
                    'Title': fieldname,
                    'Required': False,
                    'StaticName': fieldname
                }
                l = requests.post(list_url, headers=self.headers,
                                data=json.dumps(data))
            except:
                return {"error": "You messed up", "Message": l.json()}
        #return l.json()