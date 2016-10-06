'set read-only and toggle hidden attribute for a file
Const ReadOnly = 1
Const Hidden = 2

Dim objFSO, objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

'get a reference to a file
Set objFile = objFSO.GetFile("e:\data\report.doc")

'set the Readonly attribute
objFile.Attributes = objFile.Attributes Or ReadOnly

'toggle the Hidden attribute
objFile.Attributes = objFile.Attributes Xor Hidden 
