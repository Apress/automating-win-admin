Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objFolder1 = objFSO.CreateFolder("c:\data")
'create a folder below the new data folder using the Folders object’s Add method 
Set objFolder2 = objFolder1.SubFolders.Add("Word")
