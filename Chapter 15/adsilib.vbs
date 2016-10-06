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
Select Case Ucase(strType)
  Case "FTP"
    GetSiteType = "MSFTPSVC"
  Case "SMTP"
    GetSiteType = "SmtpSvc" 
  Case "NNTP"
    GetSiteType = "nntpSvc"
 Case Else
    GetSiteType = "W3SVC" 
 End Select
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
	     If nF < Cint(objSite.Name) Then nF = Cint(objSite.Name)
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

'Description
'Returns specified file or directory object
'Parameters
'strPath     File name to search for
'objIISPath  IIS object container to retrieve object from
Function GetFileDir(strPath, objIISPath)
Dim strObjectClass, objWebFileDir, objFSO

On Error Resume Next

'attempt to get the object from specified container
Set objWebFileDir = objIISPath.GetObject(strObjectClass, strPath)

'check if error occured - could not get object
If Err Then
    'create FSO object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'check if specified path is a file..
    If objFSO.FileExists(objIISPath.Path & "\" & strPath) Then
        'create the file object
        Set objWebFileDir = objIISPath.create("IIsWebFile", strPath)
    'check if specified path is a directory..
    ElseIf objFSO.FolderExists(objIISPath.Path & "\" & strPath) Then
       'create the directory object
       Set objWebFileDir = objIISPath.create("IIsWebDirectory", strPath)
    Else
        Set objWebFileDir = Nothing
    End If

End If

Set GetFileDir = objWebFileDir
Set objFSO = Nothing
End Function

'Description
'IISVersion returns the version of Internet Information Services
'installed on specified computer
'Parameters
'strComputer Name of computer to get IIS version
Function IIsVersion(strComputer)
Dim strVer, objRegistry, nRet, nValue

'create an instance of the StdRegProv registry provider
Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\default:" & "StdRegProv")

'read the MajorVersion value for IIS installed on the computer
nRet = objRegistry.GetDWordValue(, "SYSTEM\CurrentControlSet\Services\w3svc\Parameters", _
                                  "MajorVersion", nValue)

If nRet = 0 Then
    IIsVersion = CStr(nValue)
Else
    IIsVersion = "Unknown"
End If
End Function