Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("d:\data\report.doc")
objFile.Name = "newreport.doc"
