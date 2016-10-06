'compress.vbs
Option Explicit

Dim objFSO, strWMIPath, objService, objFolders, objFolder, objWMIObject

'create a FSO object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'get a reference to a WMI service on the local machine
Set objService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!root\cimv2")

'get the folder to search
Set objFolders = objFSO.GetFolder("D:\cc")

'loop through each folder
For Each objFolder In objFolders.SubFolders

    'if folder over certain size, then compress
    If objFolder.Size > 25000000 Then

     'get a reference to the directory to compress
      Set objWMIObject = _
        objService.Get("Win32_Directory.Name=" & chr(34) & _
            Replace(objFolder.Path,"\","\\") & chr(34))
    
      'invoke Compress method
      objWMIObject.Compress
    
    Wscript.Echo "Compressed folder " & objFolder.Path
    
    End If

Next

'release resources
Set objWMIObject = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
Set objService = Nothing

'\\ANARCHY\root\CIMV2:CIM_DataFile.Name="C:\\WINDOWS\\COMMAND\\EXTRACT.EXE"