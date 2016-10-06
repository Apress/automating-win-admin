Dim objMail
Set objMail = CreateObject("CDO.Send")
objMail.Profile = "My Profile" 'set the profile you want to use
If objMail.Logon() Then
 objMail.NewMessage 
 objMail.AddRecipient ("SMTP:fred@abc.com")
 objMail.Message = "Hello Fred" 
 objMail.Subject = "Message to Fred"
 objMail.Send 
 objMail.Logoff 
End iIf
