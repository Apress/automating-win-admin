Sub CreatePath(strPath)
 Dim nSlashpos, objFSO
 Set objFSO = CreateObject("Scripting.FileSystemObject")

 ' strip any trailing backslash
  If Right(strPath, 1) = "\" Then 
   strPath = Left(strPath, Len(strPath) - 1) 
  End If
   
   'if path already exists, exit
   If objFSO.FolderExists(strPath) Then Exit Sub 
    'get position of last backslash in path
    nSlashpos = InStrRev(strPath, "\")
     If nSlashpos <> 0 Then
             If Not objFSO.FolderExists(Left(strPath, nSlashpos)) Then
                    CreatePath Left(strPath, nSlashpos - 1) 
     End If
   End If

  objFSO.CreateFolder strPath
End Sub
