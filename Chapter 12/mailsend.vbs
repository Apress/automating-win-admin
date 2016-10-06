'Script: MailSend.VBS
'Description
'Sends a mail message to specified recipient
Option Explicit
Dim sExchangeProfile 'exchange profile
Dim sRecipient, sSubject, sMessage, objOneRecip, objShell
Dim objSession, objMessage 

If WScript.Arguments.Count <> 4 Then ShowUsage
     sExchangeProfile = WScript.Arguments(0)
     sRecipient = WScript.Arguments(1)
     sSubject = WScript.Arguments(2)
     sMessage = WScript.Arguments(3)
    On Error Resume Next  
    Set objSession = CreateObject("MAPI.Session")
    CheckError Err,"Unable to create mail session"

    ' logon using specified profile. 
    objSession.Logon sExchangeProfile, , False
    CheckError Err,"Unable to logon using profile: " & sExchangeProfile 

    ' create a message and fill in its properties
    Set objMessage = objSession.Outbox.Messages.Add
    objMessage.Subject = sSubject
    objMessage.Text = sMessage

    ' create the recipient
    Set objOneRecip = objMessage.Recipients.Add(,sRecipient)
    objOneRecip.Resolve   
    CheckError Err,"Unable to add recipient: " & sRecipient 

    ' send the message and log off
    objMessage.Send 
    CheckError Err,"Unable to send message " 'check for error
    objSession.DeliverNow
    objSession.Logoff
     
Sub ShowUsage()
WScript.Echo "Syntax of this script is:" &  vbCrLf & _
  "mailsend exchangeprofile, recipient, subject, message " & vbCrLf & _
  "Mailsend sends a message to a specified recipient."  & vbCrLf & _ 
  "exchangeprofile  Valid messaging profile" &  vbCrLf & _  
  "recipient        Address of recipient in format AddressType:Address. " _ 
  &  vbCrLf & _
  "subject          Message subject. " &  vbCrLf & _
  "text             Message text. " &  vbCrLf & vbCrLf & _ 
  "Example:" & vbCrLf & _
  "mailsend ""My Profile"" ""SMTP:fred@x.com"" ""message subject"" " & _
  """message text """
  WScript.Quit -1
End Sub

'Procedure: CheckError
'Description 
'Checks if error has occurred and if so, displays error information 
'and quits.
'Parameters objErr Err object
'           sMsg   Message to display if error occurrs
Sub CheckError(objErr, sMsg)
 If objErr Then
   WScript.Echo  "A fatal error has occurred: " & vbLf & sMsg & _
             vbCrLf & "Err #:" & Err _
           & vbCrLf & "Description: " & Err.Description & vbCrLf
   WScript.Quit
 End If
End Sub
