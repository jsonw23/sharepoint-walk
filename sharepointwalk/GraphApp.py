import requests

class GraphApp:

    def __init__(self, app, tenant):
        self.app = app
        self.tenant = tenant

    def fetchToken(self) -> str:
        access_token = None
        result = self.app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
        if not result:
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            access_token = result['access_token']
        else:
            raise MSALError(result)

        return access_token
    
    def fetchSite(self, relUrl):
        return self.fetchGraph(f"/sites/{self.tenant}.sharepoint.com:{relUrl}")

    def fetchGraph(self, resource: str):
        result = self.fetch(f"https://graph.microsoft.com/v1.0{resource}")
        if "error" in result:
            raise GraphError(result)
        else:
            return result

    def postGraph(self, resource: str, data=None, json=None):
        result = self.post(f"https://graph.microsoft.com/v1.0{resource}", data=data, json=json)
        if result and "error" in result:
            raise GraphError(result)
        else:
            return result

    def putGraph(self, resource: str, data=None, json=None, type=None):
        result = self.put(f"https://graph.microsoft.com/v1.0{resource}", data=data, json=json, type=type)
        if "error" in result:
            raise GraphError(result)
        else:
            return result

    def deleteGraph(self, resource: str):
        result = self.delete(f"https://graph.microsoft.com/v1.0{resource}")
        if "error" in result:
            raise GraphError(result)
        else:
            return result

    def fetchGraphPaginated(self, resource: str):
        results = self.fetchGraph(resource)
        if 'value' not in results:
            return results
        items = results['value']
        while True:
            if '@odata.nextLink' in results:
                results = self.fetch(results['@odata.nextLink'])
                items += results['value']
            else:
                break
        return items

    def fetch(self, url: str):
        access_token = self.fetchToken()
        return requests.get(url, headers={
            "Authorization": f"Bearer {access_token}"
        }).json()

    def post(self, url: str, data=None, json=None):
        access_token = self.fetchToken()
        return requests.post(url, headers={
            "Authorization": f"Bearer {access_token}"
        }, data=data, json=json)

    def put(self, url: str, data=None, json=None, type=None, headers=None):
        access_token = self.fetchToken()
        _headers = {
            "Authorization": f"Bearer {access_token}"
        }
        if type:
            _headers["Content-Type"] = type
        if headers and "Content-Range" in headers:
            _headers["Content-Range"] = headers["Content-Range"]
        response = requests.put(url, headers=_headers, data=data, json=json)
        if response.ok:
            return response.json()
        else:
            raise Exception(response.json())

    def delete(self, url: str):
        access_token = self.fetchToken()
        _headers = {
            "Authorization": f"Bearer {access_token}"
        }
        response = requests.delete(url, headers=_headers)
        if response.ok:
            return response
        else:
            raise Exception(response.json())



class MSALError(Exception):
    pass

class GraphError(Exception):
    pass