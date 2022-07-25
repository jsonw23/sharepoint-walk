from .GraphApp import GraphApp, GraphError
from .DriveItem import DriveItem, File, Folder


class SharePointWalker:

    def __init__(self, app: GraphApp):
        self.app = app

    def walk(self, location: str):
        (siteID, driveID, item, children) = self.__findLocation(location)
        
        # item['path'] is /drives/{drive-id}/root:/{path-relative-to-root}:/children
        # mimick os.walk
        # yield return tuples with:
        # - root: the path to the item, between the ':'s in the item['path']
        # - folders: list of driveitems that are folders
        # - files: list of driveitems that are files
        rootPath: str = item['path'].split(':')[1][1:]
        rootFolder = Folder(item)
        folderStack = []
        while True:
            folders = []
            files = []
            for child in children:
                driveItem = DriveItem.encapsulate(child)
                if isinstance(driveItem, Folder):
                    folders.append(driveItem)
                if isinstance(driveItem, File):
                    files.append(driveItem)

            yield (rootFolder, folders, files)

            if len(folders):
                folderStack.append(folders[::-1])
                rootStack = rootPath.split('/')
            
            if not len(folderStack):
                # no more to scan
                return

            nextFolder: DriveItem = folderStack[-1].pop()
            if not len(folderStack[-1]):
                folderStack.pop()
            rootPath = nextFolder.path

            try:
                children = self.app.fetchGraphPaginated(f"/drives/{driveID}/root:{rootPath}:/children")
                if len(children):
                    rootFolder = Folder(children[0]["parentReference"])
            except GraphError as e:
                children = []


    def __findLocation(self, location: str):
        # /sites/SiteName:DocLibrary/path/to/start
        [siteUrl, path] = location.split(":")
        pathSegs = path.split("/")
        docLibrary = pathSegs[0]
        pathSegs = pathSegs[1:]
        
        # find the site
        site = self.app.fetchSite(siteUrl)
        
        # find the drive
        drive = None
        if "id" in site:
            drives = self.app.fetchGraph(f"/sites/{site['id']}/drives")
            for _drive in drives['value']:
                if _drive['name'] == docLibrary:
                    drive = _drive

        # find the item
        parent = None
        children = None
        if "id" in drive:
            children = self.app.fetchGraphPaginated(f"/drives/{drive['id']}/root:/{'/'.join(pathSegs)}:/children")
            if len(children):
                # use the parent of the first item
                parent = children[0]['parentReference']
        
        return (site['id'], drive['id'], parent, children)