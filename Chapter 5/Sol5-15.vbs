Dim objFSO, objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
 objFSO.DeleteFile "d:\data\report.doc"

Set objFile = objFSO.GetFile("d:\data\payroll.xls")
objFile.Delete
