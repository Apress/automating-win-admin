'Script: ISAlive.VBS
'Description
'Checks if machines are unavailable and E-maile-mails a specified user
'if can't reach machines (machines might be down)
Dim objPing, strResult, objMail
Set objPing = CreateObject("SScripting.IPNetwork")

strResult = ""
Ping "Mars", strResult
Ping "Jupiter", strResult
Ping "Thor", strResult

'check if the return result is not empty - if not empty, then unable to ping
'one or more machines
If strResult <> "" Then
 sResult = "Unable to contact following machines:" & vbCrLf & sResult
 Set objMail = CreateObject("CDO.Send")
 objMail.Profile = "SCB"
 objMail.Logon
 objMail.NewMessage
 objMail.AddRecipient "SMTP:administrator@c3i.com"
 objMail.Message = "Machine(s) Not Available "
 objMail.Subject = strResult
 objMail.Send
 objMail.Logoff
End If

Sub Ping(strHost, ByRef strMsg)
  If Not objPing.Ping(strHost) = 0 Then strMsg = strMsg & strHost & vbCrLf
End Sub
