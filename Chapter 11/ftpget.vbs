'ftpget.vbs
'copies files from FTP directory to local directory
Const SynchronousMode = 1
Dim objFTP, nF, strName
 Set objFTP = CreateObject("Mabry.FtpXObj")
'connect to FTP server Thor
 objFTP.Blocking =  SynchronousMode 
 objFTP.LogonName = "administrator"  
objFTP.LogonPassword = "downunder" 
 objFTP.host = "thor" 'hostname
 objFTP.Connect

 If Err Then
   WScript.Echo "Error connecting to FTP host" 
 End If

  'get a directory listing for remote machine
  objFTP.GetDirList "/" 
  For nF = 0 To objFTP.DirItems - 1
    'get item from array
    strName = objFTP.DirItem(nf)
    
    'check if item is not a directory
    If Not InStr(strName, "<dir>") > 0 Then
     'get the name of the file, which for IIS starts at position 40
     strName = Mid(strName, 40)
     'strip off carriage return/line feed from end of string
     strName = Left(strName, Len(strName) - 2)
     'get file, store in local drive
     objFTP.GetFile strSrcDir & strName, "d:\data\" & strName
    End If
  Next
  objFTP.Disconnect
