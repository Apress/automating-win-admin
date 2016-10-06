Dim objFSO, objFolder
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Data")
'copy items from folder to network folder
objFolder.Copy"H:\Data"
