'iniwrite.vbs
'updates an INI file entry
Const ForWriting = 2 
Const ForAppending = 8

UpdateINI "e:\settings.ini", "DateLastAccessed", Now

Sub UpdateINI(strINIFile, strKey, strValue)
 Dim objFSO, objTextFile, strFileText, aDataFile, nF, bFound

 Set objNetwork = CreateObject("WScript.Network")

 Set objFSO = CreateObject("Scripting.FileSystemObject")

 'open the user file
 Set objTextFile = objFSO.OpenTextFile(strINIFile)

 'read the whole file
 If Not objTextFile.AtEndOfStream Then
   strFileText = objTextFile.ReadAll

   'split the file into an array.
   aDataFile = Split(strFileText, vbCrLf)

   'loop through each item in the array
   For nF = 0 to Ubound(aDataFile)
      If Left(aDataFile(nF), Len(strKey)+1)=strKey & "="  Then
       aDataFile(nF) = strKey & "=" & strValue 
       bFound = True
       Exit For
    End If
   Next

 End If

 objTextFile.Close
 
 'if entry was found then write back contents
 If bFound Then
   strFileText = Join(aDataFile, vbCrLf)
   Set objTextFile = _
           objFSO.OpenTextFile(strINIFile, ForWriting )
   objTextFile.Write strFileText
 Else
   'entry not found, add new entry to end of file
   Set objTextFile = _
           objFSO.OpenTextFile(strINIFile, ForAppending )
   objTextFile.WriteLine strKey & "=" & strValue 
 End If

 objTextFile.Close
End Sub
