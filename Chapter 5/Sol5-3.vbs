Dim objFSO, objFolder, objSub, nTotal
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("D:\Users")

nTotal = 0
'loop through each subfolder, displaying its size 
For each objSub In objFolder.SubFolders
   Wscript.Echo "Folder " & objSub.Name & " is " & objSub.Size _
   & " bytes"
    nTotal = nTotal + objSub.Size
Next

Wscript.Echo "Total for all folders:" & nTotal & " bytes"
