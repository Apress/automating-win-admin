'readini.vbs
Dim dAccess

dAccess = ReadINI("e:\data.ini", "DateLastAccessed")

Function ReadINI(strINIFile, strKey)
 Dim objFSO, objTextFile, strLine

 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objTextFile = objFSO.OpenTextFile(strINIFile)

 'loop through each line of file and check for key value
 Do While Not objTextFile.AtEndOfStream
    strLine = objTextFile.ReadLine
    If Left(strLine, Len(strKey) + 1) = strKey & "=" Then
        ReadINI = Mid(strLine, InStr(strLine, "=") + 1)
    End If
 Loop

 objTextFile.Close

End Function
