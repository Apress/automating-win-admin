'deltemp.vbs
Dim objDelTempFiles , objFSO

'create an instance of the ENTWSH.RecurseDir object.
'any sinked events will be prefixed with ev_
Set objDelTempFiles = Wscript.CreateObject ("ENTWSH.RecurseDir","ev_")

Set objFSO = CreateObject("Scripting.FileSystemObject")

objDelTempFiles.Path = "d:\data" 'set the path to search
'set the filter for files to find. This will find all files with tmp extension
objDelTempFiles.Filter = "\.tmp$" 

objDelTempFiles.Process

Sub ev_FoundFile(strPath)
      objFSO.DeleteFile strPath 
End Sub
