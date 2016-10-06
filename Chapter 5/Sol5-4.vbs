Dim objFSO, objFile, nSize

'create an instance of an FSO object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'get a reference to a specified file
Set objFile = objFSO.GetFile("d:\data\report.doc")

nSize = objFile.Size 'get the size of the file
