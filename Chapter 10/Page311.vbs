Dim objService, objWMIFiles, nResult, objFile

On Error Resume Next

'get a reference to a WMI service
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}")

'return all zip files from the D: drive that are compressed
Set objWMIFiles = _
        objService.ExecQuery("Select * From CIM_DataFILE Where Drive='D:'" & _ 
                                           " And Extension='zip' And Compressed=True")
'loop through any compressed files and uncompress them
For Each objFile In objWMIFiles
    WScript.Echo "Uncompressing file " & objFile.Name
    objFile.Uncompress
Next
