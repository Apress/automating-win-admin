'adsilib.vbs
'Description: Contains routines used by ADSI scripts

'Gets the value of a server object based on it's server comment/name
'Parameters:
'objWebService  WebService object
'strSiteName    Site name you wish to get value
'Returns: Site number, blank string if not found
Function FindSiteNumber(objWebService, strSiteName, strType)
Dim nF, objSite

 nF = ""

 'loop through each site, find available site #
 For Each objSite In objWebService
    'check if the object is a web site
    
    If strcomp(objSite.Class,"IIs" & strType & "Server",1)=0 Then

     'check if server comment is same as specified server name
     If Ucase(objSite.ServerComment) = Ucase(strSiteName) Then 
	nF = objSite.Name
	Exit For
     End If 	
    End If
 Next
FindSiteNumber = nF
End Function

'returns the 
'Parameters
'strType     site type, web or FTP
Function GetSiteType(strType)

Dim strSvc

Select Case Ucase(strType)
  Case "FTP"
    strSvc = "MSFTPSVC"
  Case "SMTP"
    strSvc = "SmtpSvc" 
  Case "NNTP"
    strSvc = "nntpSvc"
 Case Else
    strSvc = "W3SVC" 

End Select

GetSiteType = strSvc

End Function 

'Find next available site number
'
Function FindNextSite(objService)
Dim nF, objSite

nF = 0

'loop through each site, find available site #

For Each objSite In objService
    'check if object is a IIS site
    If Left(objSite.Class,3) = "IIs" And _
		Right(objSite.Class,6)= "Server" Then
     If nF < objSite.Name Then nF = objSite.Name
    End If
    
Next

FindNextSite = nF + 1

End Function

'check if script is being run interactively
'Returns:True if run from command line, otherwise false
Function IsCscript()
 If strcomp(Right(Wscript.Fullname,11),"cscript.exe",1)=0 Then 
   IsCscript = True
 Else
   IsCscript = False
 End If
End Function

'display an error message and exist script
'Parameters:
'strMsg Message to display
Sub ExitScript(strMsg)
 Wscript.Echo strMsg
 Wscript.Quit(-1)
End Sub