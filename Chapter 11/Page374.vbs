Const MSS_NOT_LOGGED_ON = 1
Const MSS_LOGGED_ON = 0

Dim objMessenger
Set objMessenger = CreateObject("Messenger.msgrobject")

'if not logged on then logon
If objMessenger.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
  Set objMessenger = StartMessenger("Freds@hotmail.com", "sderf!@#")
 
   'wait until successfully logged on 
   Do While Not objMessenger.Services.PrimaryService.Status = MSS_LOGGED_ON
     WScript.Sleep 100
   Loop
End If
