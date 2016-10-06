Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFile = objFSO.GetFile("C:\Data\report.doc")
'copy items from folder to network folder
objFile.Copy "H:\Data\"
