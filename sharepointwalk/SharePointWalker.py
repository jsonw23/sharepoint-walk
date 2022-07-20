import requests

class SharePointWalker:
    app = None
    tenant = None

    def __init__(self, app, tenant):
        self.app = app
        self.tenant = tenant

    def walk(self, location: str):
        (siteID, driveID, item, children) = self.__findLocation(location)
        
        # item['path'] is /drives/{drive-id}/root:/{path-relative-to-root}:/children
        # mimick os.walk
        # yield return tuples with:
        # - root: the path to the item, between the ':'s in the item['path']
        # - folders: list of driveitems that are folders
        # - files: list of driveitems that are files
        rootPath: str = item['path'].split(':')[1][1:]
        folderStack = []
        while True:
            folders = []
            files = []
            for child in children:
                if 'folder' in child:
                    folders.append(child)
                if 'file' in child:
                    files.append(child)

            yield (rootPath, folders, files)

            if len(folders):
                folderStack.append(folders[::-1])
                rootStack = rootPath.split('/')
            
            nextFolder = folderStack[-1].pop()
            if not len(folderStack[-1]):
                folderStack.pop()
            rootPath = nextFolder['parentReference']['path'].split(':')[1][1:] + '/' + nextFolder['name']

            children = self.__fetchGraphPaginated(f"/drives/{driveID}/root:/{rootPath}:/children")


    def __findLocation(self, location: str):
        # /sites/SiteName:DocLibrary/path/to/start
        [siteUrl, path] = location.split(":")
        pathSegs = path.split("/")
        docLibrary = pathSegs[0]
        pathSegs = pathSegs[1:]
        
        # find the site
        site = self.__fetchSite(siteUrl)
        
        # find the drive
        drive = None
        if "id" in site:
            drives = self.__fetchGraph(f"/sites/{site['id']}/drives")
            for _drive in drives['value']:
                if _drive['name'] == docLibrary:
                    drive = _drive

        # find the item
        parent = None
        children = None
        if "id" in drive:
            children = self.__fetchGraphPaginated(f"/drives/{drive['id']}/root:/{'/'.join(pathSegs)}:/children")
            if len(children):
                # use the parent of the first item
                parent = children[0]['parentReference']
        
        return (site['id'], drive['id'], parent, children)

    def __fetchToken(self) -> str:
        access_token = None
        result = self.app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
        if not result:
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            access_token = result['access_token']
        else:
            raise MSALError(result)

        return access_token
    
    def __fetchSite(self, relUrl):
        return self.__fetchGraph(f"/sites/{self.tenant}.sharepoint.com:{relUrl}")

    def __fetchGraph(self, resource: str):
        result = self.__fetch(f"https://graph.microsoft.com/v1.0{resource}")
        if "error" in result:
            raise GraphError(result)
        else:
            return result

    def __fetchGraphPaginated(self, resource: str):
        results = self.__fetchGraph(resource)
        items = results['value']
        while True:
            if '@odata.nextLink' in results:
                results = self.__fetch(results['@odata.nextLink'])
                items += results['value']
            else:
                break
        return items

    def __fetch(self, url: str):
        access_token = self.__fetchToken()
        return requests.get(url, headers={
            "Authorization": f"Bearer {access_token}"
        }).json()



class MSALError(Exception):
    pass

class GraphError(Exception):
    pass