Const FTP_TRANSFER_TYPE_ASCII = 1
Const FTP_TRANSFER_TYPE_BINARY = 2
Dim objFTP

Set objFTP = CreateObject("AspInet.FTP")

If objFTP.FTPGetFile("ftp.acme.com", "userid", "password", _
      "/data.txt", "d:\data\data.txt", True, FTP_TRANSFER_TYPE_ASCII) Then
     WScript.Echo "File download succeeded"
End If

If objFTP.FTPPutFile("ftp.acme.com", "userid", "password", _
      "/data.txt", "d:\data\data.txt", FTP_TRANSFER_TYPE_ASCII) Then
     Wscript.Echo "File download succeeded"
End If
