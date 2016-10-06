'Description
'Returns specified file or directory object
'Parameters
'strPath     File name to search for
'objIISPath  IIS object container to retrieve object from
Function GetFileDir(strPath, objIISPath)
Dim strObjectClass, objWebFileDir, objFSO

On Error Resume Next

'attempt to get the object from specified container
Set objWebFileDir = GetObject(objIISPath.ADsPath & "/" & strPath)

'check if error occured - could not get object
If Err Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'check if specified path is a file..
    If objFSO.FileExists(objIISPath.Path & "\" & strPath) Then
        'create the file object
     Set objWebFileDir = objIISPath.create("IIsWebFile", strPath)
     objWebFileDir.SetInfo
    'check if specified path is a directory..
    ElseIf objFSO.FolderExists(objIISPath.Path & "\" & strPath) Then
       'create the directory object
       Set objWebFileDir = objIISPath.create("IIsWebDirectory", strPath)
      objWebFileDir.SetInfo
    Else
        Set objWebFileDir = Nothing
    End If

End If

Set GetFileDir = objWebFileDir
End Function
