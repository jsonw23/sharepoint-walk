import mimetypes
import requests
import os

from sharepointwalk.GraphApp import GraphApp

class DriveItem:

    def __init__(self, driveItem):
        self.__driveItem = driveItem

    @property
    def name(self) -> str:
        if 'name' in self.__driveItem:
            return self.__driveItem['name']
        else:
            return self.path.split("/")[-1]
    
    @property
    def path(self) -> str:
        return self.__driveItem['parentReference']['path'].split(':')[1] + '/' + self.name

    @property
    def id(self) -> str:
        return self.__driveItem['id']
    
    @property
    def driveID(self) -> str:
        if 'parentReference' in self.__driveItem:
            return self.__driveItem['parentReference']['driveId']
        else:
            return self.__driveItem['driveId']

    @property
    def parentID(self) -> str:
        if 'parentReference' in self.__driveItem:
            return self.__driveItem['parentReference']['id']
        else:
            return None

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

    @property
    def size(self) -> int:
        return self._DriveItem__driveItem["size"]
    
    def download(self, to="", app: GraphApp=None) -> str:
        path = os.path.join(to, self.name)
        r = requests.get(self._DriveItem__driveItem["@microsoft.graph.downloadUrl"])
        if r.ok:
            with open(path, "wb") as dl:
                dl.write(r.content)
            return path
        elif r.status_code == 401:
            # download url is expired
            if app:
                result = app.fetchGraph(f"/drives/{self.driveID}/items/{self.id}")
                if "@microsoft.graph.downloadUrl" in result:
                    r = requests.get(result["@microsoft.graph.downloadUrl"])
                    if r.ok:
                        with open(path, "wb") as dl:
                            dl.write(r.content)
                        return path

class Folder(DriveItem):
    
    @property
    def path(self) -> str:
        if 'parentReference' in self._DriveItem__driveItem:
            return self._DriveItem__driveItem['parentReference']['path'].split(':')[1] + '/' + self.name
        else:
            return self._DriveItem__driveItem['path'].split(':')[1]

def newFolder(app: GraphApp, driveID: str, parentID: str, name: str) -> Folder:
    result = app.postGraph(f"/drives/{driveID}/items/{parentID}/children", json={
        "name": name,
        "folder": { },
        "@microsoft.graph.conflictBehavior": "rename"
    })
    return Folder(result)

def uploadFile(app: GraphApp, parentFolder: Folder, path: str) -> File:
    (type, _) = mimetypes.guess_type(path)
    info = os.stat(path)
    sizeMb = info.st_size / (1024 * 1024)
    if sizeMb > 4:
        return uploadLargeFile(app, parentFolder, path, info.st_size)
    else:
        with open(path, "rb") as f:
            name = os.path.basename(path)
            result = app.putGraph(f"/drives/{parentFolder.driveID}/items/{parentFolder.id}:/{name}:/content", data=f.read(), type=type)
            return File(result)

def uploadLargeFile(app: GraphApp, parent: Folder, path: str, size: int) -> File:
    # create an upload session
    session = app.postGraph(f"/drives/{parent.driveID}/items/{parent.id}:/{os.path.basename(path)}:/createUploadSession", json={
        "item": {
            "@microsoft.graph.conflictBehavior": "replace"
        }
    })
    
    # figure out what size chunk
    chunkFactor = 320 * 1024
    maxChunk = 60 * 1024 * 1024
    chunkSize = 10 * 1024 * 1024 # for now just 10mb

    with open(path, "rb") as f:
        bytesSent = 0
        while bytesSent < size:
            chunk = f.read(chunkSize)
            response = app.put(session["uploadUrl"], data=chunk, headers={
                "Content-Range": f"bytes {bytesSent}-{bytesSent + len(chunk) - 1}/{size}"
            })
            bytesSent += len(chunk)
        return File(response)
