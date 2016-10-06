Const MSS_NOT_LOGGED_ON = 1
Const MSS_LOGGED_ON =0
Const MSTATE_ONLINE = 2
Const MLIST_CONTACT =0
Const MLIST_ALLOW = 1
Const MLIST_BLOCK = 2
Const MLIST_REVERSE = 3

Const MMSGTYPE_NO_RESULT = 0
Const MMSGTYPE_ERRORS_ONLY =1
Const MMSGTYPE_ALL_RESULTS = 2

Const MSGR_E_FAIL = -2147467259 
Const MSGR_S_OK = 0

Dim MsgHeader
MsgHeader = "Mime-Version: 1.0" & vbCrLf & _
            "Content-Type: text/plain; charset=UTF-8" & vbCrLf & vbCrLf

'creates a MSN messenger object and logs in if
'not already connected
'Parameters:
'strUserID     User ID to connect
'strPassword   password to connect with
'Returns
'Messenger object if succesfull, Nothing if not
'successful
Function StartMessenger(strUserID, strPassword)
  Dim objMessenger
  Set objMessenger = CreateObject("Messenger.Msgrobject")

  'if user not logged on then logon
  If objMessenger.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
    objMessenger.Logon strUserID, strPassword, _
                       objMessenger.Services.PrimaryService

   'wait until successfully logged on 
   Do While Not objMessenger.Services.PrimaryService.Status = MSS_LOGGED_ON
    'if primary service is logged off then logon attempt unsuccessful
     If objMessenger.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
	    StartMessenger = Nothing
        Exit Function 
     End If
    Wscript.Sleep 100
   Loop
  End If
  
  Set StartMessenger = objMessenger
End Function