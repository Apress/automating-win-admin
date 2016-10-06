'compvers.vbs
Dim objFSO, strFolder1, strFolder2, objFolder
Dim objFile, strVer1

'check if correct number of arguments passed
 If WScript.Arguments.Count <> 2 Then 
   WScript.Echo  _
     WScript.ScriptName & _
     " compares versions of files in two folders." & vbCrLf & _ 
     "Output is all files which exist in both folders but " & vbCrLf & _
     "have different version stamps." & vbCrLf & _
     "Syntax:" &  vbCrLf & _
     WScript.ScriptName & " folder1 folder2" &  vbCrLf  & vbCrLf
      WScript.Quit -1
  End If

  strFolder1 = WScript.Arguments(0) 
  strFolder2 = WScript.Arguments(1)

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  EnsureFolder strFolder1 
  EnsureFolder strFolder2

  Set objFolder = objFSO.GetFolder(strFolder1)
 
 'loop through each file in folder and compare with second folder
 For Each objFile In objFolder.Files
   strPath2 = strFolder2 & objFile.Name
   If objFSO.FileExists(strPath2) Then
      CompareFiles objFile, objFSO.GetFile(strPath2)
   End If
 Next

'check if folder exists
Sub EnsureFolder(strFolder)
    If Not objFSO.FolderExists(strFolder) Then
        WScript.Echo strFolder & " is not a valid folder." & vbCrLf
        WScript.Quit -1
    End If

  'append backslash to folder if does not exist
  If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
End Sub

'compare versions of two passed file objects and display version info
'if different
Sub CompareFiles(objFile1, objFile2)

    If  objFSO.GetFileVersion(objFile1.Path) = _
           objFSO.GetFileVersion(objFile2.Path) Then Exit Sub

     WScript.Echo objFile1.Path & " has version [" & _
                   strVer1 & "] and " & _
                  objFile2.Path & " has version [" & _
                   strVer2 & "]"
End Sub
