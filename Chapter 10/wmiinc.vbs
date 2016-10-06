'wmiinc.vbs
'contains reusable WMI support code 
Option Explicit

Const WMIConst = "winmgmts:{impersonationLevel=impersonate}!"
Const wbemCimtypeDatetime =101
Const wbemCimtypeString = 8
Const wbemCimtypeChar16 = 103

'ExitScript
'Displays message and terminates script
'Parameters:
'strMessage    Message to display
'bStdOut       Boolean value. If true then writes to StdErr
Sub ExitScript(strMessage, bStdOut)

If bStdOut Then
    'get a reference to the StdErr object stream
    Wscript.StdErr.WriteLine strMessage
Else
    Wscript.Echo strMessage
End If
    Wscript.Quit
End Sub


Function Convert2DMTFDate(dDate, nTimeZone)
Dim sTemp, sTimeZone

sTimeZone = nTimeZone
If nTimeZone>=0 Then sTimeZone = "+" & sTimeZone

sTemp = Year(Now) & Pad(Month(dDate), 2, "0") & Pad(Day(dDate), 2, "0")
sTemp = sTemp & Pad(Hour(dDate), 2, "0") & Pad(Minute(dDate), 2, "0")
sTemp = sTemp & Pad(Second(dDate), 2, "0") & ".000000" & sTimeZone


Convert2DMTFDate = sTemp
End Function

Function Pad(sPadString, nWidth, sPadChar)

    If Len(sPadString) < nWidth Then
        Pad = String(nWidth - Len(sPadString), sPadChar) & sPadString
    Else
        Pad = sPadString
    End If

End Function

'DMTFDate2String
'Converts WMI DMTF dates to a readable string
'Parameters:
'strDate    Date in DMTF format
'Returns
'formatted date string
Function DMTFDate2String(strDate)
 strDate = Cstr(strDate)
 DMTFDate2String = Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2) _
          & "/" & Mid(strDate, 1, 4) & " " & Mid(strDate, 9, 2) _
          & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
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


'GetBinarySID
'Returns user's SID as array of integer values
'Parameters:
'strAccount User account in Domain\username or username format
'Returns
'array of integer values if successful, otherwise Null
Function GetBinarySID(strAccount)

Dim objAccounts, objAccount, bDomain, bFound, objSIDAccount

bDomain = False
bFound = False

'check if a backslash exists in the account name.. if so search for
'domainname\accountname
If InStr(strAccount, "\") > 0 Then bDomain = False

'get an instance of the Win32_Account object
Set objAccounts = GetObject( WMICONST & "root\cimv2") _
                    .InstancesOf("Win32_Account")

'loop through each account
For Each objAccount In objAccounts
    'if domain name specified, search against account caption
    If bDomain Then
        'check if name is found
        If StrComp(objAccount.Caption, strAccount, 1) = 0 Then
            bFound = True
            Exit For
        End If
    Else 'check against just user name
        If StrComp(objAccount.Name, strAccount, 1) = 0 Then
        'check if name is found
            bFound = True
            Exit For
        End If
    End If
    
Next

'if found then retrieve SID binary array
If bFound Then

Set objSIDAccount=GetObject(WMICONST & "Win32_SID.SID=" _
                & Chr(34) & objAccount.sid & Chr(34))

    GetBinarySID = objSIDAccount.BinaryRepresentation
Else
    GetBinarySID = Null
End If

End Function

Class WMISupport

 Dim objLocator, strErrorMsg, objService
 Dim strServer, strNameSpace, strUserName, strPassword

 Private Sub Class_Initialize()
  objService = Empty
  strServer = Empty 
  strNameSpace = Empty 
  strUserName = Empty  
  strPassword = Empty  
 End Sub

 Private Sub Class_Terminate()
  Set objLocator = Nothing
  Set objService = Nothing
 End Sub

'creates WMI session and returns WMI service object
 Function Connect()
    On Error Resume Next

    'set the default return value
    Connect = Null

    'create locator object
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
        
   If Err Then 
    strErrorMsg =  "Error getting reference to WBEM locator object"
    Exit Function
   End If    

   'connect to specified machine
   Set objService = objLocator.ConnectServer(strServer, _
                             strNameSpace, strUserName, _
                             strPassword)

   If Err Then 
    strErrorMsg =  "Error connecting to " & strServer
    Exit Function
   End If    
    'set impersonation level to Impersonate
    objService.Security_.ImpersonationLevel = 3
    
   If Err Then 
    strErrorMsg =  "Error setting security level" 
    Exit Function
   End If    

    Set Connect = objService
 End Function

 Public Property Let Computer (strComputerName)
    strServer = strComputerName
 End Property 

 Public Property Let UserName (strUser)
    strUserName =strUser 
 End Property 

 Public Property Let Password (strPass)
    strPassword = strPass
 End Property 

 Public Property Let NameSpace (strNameSpc)
    strNameSpace = strNameSpc
 End Property 

 Public Property Get ErrorMessage()
  Set ErrorMessage = strErrorMsg
 End Property

End Class
