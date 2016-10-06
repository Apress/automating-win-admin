Dim objFSO, objFolder
Set objFSO = CreateObject("Scripting.FileSystemObject")
'delete the folder e:\data\word
objFSO .DeleteFolder "e:\data\word"

'get and delete the folder e:\data\excel
Set objFolder = objFSO.GetFolder("e:\data\Excel")
objFolder.Delete
