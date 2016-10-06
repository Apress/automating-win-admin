 'ftpxcopy.vbs
  'copies directory and all sub directories to FTP server
  Const SynchronousMode = 1
  Dim objFile, objFSO, objFTP, strUser
  Dim strSrcRoot, nStart, strDstRoot
 
If Not WScript.Arguments.Count = 5 Then
    ShowUsage
     Wscript.Quit
  End If

  'get user id, host, password and source/destination directories
   strSrcRoot = Wscript.Arguments(3)
   strDstRoot = Wscript.Arguments(4)
   If Not Right(strSrcRoot,1) = "\" Then strSrcRoot = strSrcRoot & "\"
   If strDstRoot = "/" Then strDstRoot = ""

   nStart = Len(strSrcRoot)
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objEvent = Wscript.CreateObject ("ENTWSH.RecurseDir","ev_")
   Set objFTP = CreateObject("Mabry.FtpXObj")
 
   objFTP.Blocking = SynchronousMode 
   objFTP.host = Wscript.Arguments(0)
   objFTP.Connect Wscript.Arguments(1), Wscript.Arguments(2)

   objEvent.Path = strSrcRoot
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Call objEvent.Process()
   objFTP.Disconnect
   
 Sub ShowUsage
WScript.Echo "ftpxcopy.vbs copies local directory to FTP server" _
    & vbLf & "Syntax:" &  vbCrLf & _
   "ftpxcopy host user password source destination " & vbCrLf & _
   "host        FTP server to copy to" & vbCrLf & _ 
   "user        user name to log on to FTP server" & vbCrLf & _ 
   "password    password to logon onto FTP server" & vbCrLf & _
   "source      local path to source directory" & vbCrLf & _
   "destination path to FTP directory" & vbCrLf & _
   "Example: ftpxcopy acme freds sderf d:\data\website\acme /webroot/acme"
 End Sub
 
 Sub ev_FoundFile(strPath)
   On Error Resume Next 
   'get a reference to specifed file
   Set objFile = objFSO.GetFile(strPath)

   'convert file path to corresponding FTP directory   
   strFTPDir = Mid(objFile.ParentFolder, nStart)
   strFTPDir = strDstRoot & Replace(strFTPDir, "\","/")

   If Not Right(strFTPDir,1) = "/" Then strFTPDir =  strFTPDir & "/"
   objFTP.PutFile strPath, strFTPDir & objFile.Name
  
   If Not objFTP.LastError = 0 Then 
     MakeDirPath (Left(strFTPDir,Len(strFTPDir)-1))   
    objFTP.PutFile strPath, strFTPDir & "/" &  objFile.Name
   End If
End Sub

'Procedure MakeDirPath
'Description
'Creates a directory path on remote FTP server
'Parameters
'strPath FTP directory path to create
Sub MakeDirPath(strPath)
Dim nF, strRest, strNextPath
On Error Resume Next
bDone = False
strNextPath = strPath

 Do While Not bDone
 'check if directory exists
  objFTP.ChangeDir strNextPath ', strRest
 'if directory doesn't exist, then parse next level in path
  If Not objFTP.LastError = 0 Then 
   nF = InStrRev(strNextPath, "/")
     strRest = Mid(strNextPath, nF) & strRest
     strNextPath = Left(strNextPath, nF - 1)
   Else
    'directory found, create path below it
    strRest = Mid(strRest,2) & "/"
     nF = 0
     Do While True
        nF = Instr(strRest,"/")
       strNextPath = strNextPath & "/" & Left(strRest, nF - 1)
        objFTP.CreateDir strNextPath
                    WScript.Echo "Creating directory " & strNextPath
        If nF = Len(strRest) Then Exit Do
       strRest = Mid(strRest, nF + 1)
     Loop 
    bDone = True
   End If
 Loop
End Sub
