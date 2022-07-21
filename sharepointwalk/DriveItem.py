import requests
import os

from sharepointwalk.GraphApp import GraphApp

class DriveItem:

    def __init__(self, driveItem):
        self.__driveItem = driveItem

    @property
    def name(self) -> str:
        return self.__driveItem['name']
    
    @property
    def path(self) -> str:
        return self.__driveItem['parentReference']['path'].split(':')[1] + '/' + self.name
    
    @property
    def driveID(self) -> str:
        return self.__driveItem['parentReference']['driveId']

    @property
    def parentID(self) -> str:
        return self.__driveItem['parentReference']['id']

    def encapsulate(driveItem):
        if "folder" in driveItem:
            return Folder(driveItem)
        elif "file" in driveItem:
            return File(driveItem)

    def __str__(self) -> str:
        return self.name

    def __repr__(self) -> str:
        return str(self)
    
class File(DriveItem):
    
    def download(self, to="") -> str:
        path = os.path.join(to, self.name)
        r = requests.get(self._DriveItem__driveItem["@microsoft.graph.downloadUrl"])
        with open(path, "wb") as dl:
            dl.write(r.content)
        return path

class Folder(DriveItem):
    pass

def newFolder(app: GraphApp, driveID: str, parentID: str, name: str) -> Folder:
    result = app.postGraph(f"/drives/{driveID}/items/{parentID}/children", json={
        "name": name,
        "folder": { },
        "@microsoft.graph.conflictBehavior": "rename"
    })
    return Folder(result)