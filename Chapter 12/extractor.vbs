'extractor.vbs
Dim objMessage, objSession, objMessages, objFilter
Dim objFolder, sline, sBody 
Dim nLast, nPos, nF

Set objSession = CreateObject("MAPI.Session")

'logon using an existing session..
objSession.Logon , , False, False
Set objShell = CreateObject("WScript.Shell")
'get the command line arguments
On Error Resume Next
nLast = 1
'get the message..
If WScript.Arguments.Count = 0 Then 
  WScript.Echo "Requires MAPI message ID" 
  WScript.Quit
End If

Set objMessage = objSession.GetMessage(WScript.Arguments(0))

'check if valid message
If objMessage Is Nothing Then WScript.Quit
   
 'get the body of the message
 sBody = objMessage.Text

'check body is not empty..
If Len(sBody)=0 Then WScript.Quit
 'loop through and process each line of the message
 Do
  'get the end of the current line
   nPos = InStr(nLast, sBody, vbCrLf)
    If nPos = 0 Then nPos = Len(sBody)
     sline = Trim(Mid(sBody, nLast, nPos - nLast))
    'check the first 4 characters of each line
    Select Case UCase(Left(sline, 4))
    
    Case "XTR:" 'extract command
     'get the position of a comma in the line -
     ' the text after the comma is the directory to extract to
     nF = InStr(sline, ",")
     If Not nF Then
       ProcessAttachments objMessage, "XTR", _ 
                 Trim(Mid(sline, 5, nF - 5)), Trim(Mid(sline, nF + 1))
     End If
    
    Case "EXE:" 'execute command
    ProcessAttachments objMessage, "EXE", Trim(Mid(sline, 5)), ""
     
    Case "DEL:" 'delete command
     ProcessAttachments objMessage, "DEL", Mid(sline, 5), ""
    
    End Select
    nLast = nPos + 2
  Loop While nLast < Len(sBody)
objSession.Logoff

Function ProcessAttachments(objMessage, sType, sFile, sPath)
Dim objAttachment, objFS, objFolder, objShell, sTemp
    For Each objAttachment In objMessage.Attachments
     'check if the current attachment name is equal to the one you
     'want to process
     If StrComp(objAttachment.Name, sFile, vbTextCompare) = 0 Then
                Select Case sType
            Case "EXE"
             Set objShell =CreateObject("WScript.Shell")
             sTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
             objAttachment.WriteToFile sTemp & "\" & sFile
               objShell.Run sTemp & "\" & sFile, 1, True
                Set objFS = CreateObject("Scripting.FileSystemObject")
         Set objFile = objFS.GetFile(sTemp & "\" & sFile)
         objFile.Delete 
          Case "XTR"
                'create a file system object
                Set objFS = CreateObject("Scripting.FileSystemObject")
                Set objShell =CreateObject("WScript.Shell")
              sPath = objShell.ExpandEnvironmentStrings(sPath)
 
                'if folder doesn't exist, exit function
                If Not objFS.FolderExists(sPath)Then
                  Exit Function
                End If
                'if folder exists, then extract attachment into folder
                 objAttachment.WriteToFile sPath & "\" & sFile
               
            Case "DEL"
                 Set objShell = CreateObject("WScript.Shell")
                sPath = objShell.ExpandEnvironmentStrings(sPath)
                'create a file system object
                Set objFS = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFS.GetFile(sPath)
        objFile.Delete 
       End Select          
         Exit For
    End If   
    Next
End Function
