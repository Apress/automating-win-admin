Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "administrator@acme.com"
objMessage.From = "joeb@acme.com" 
objMessage.To = "test@acme.com" 
objMessage.CC = "test@acme.com" 
objMessage.BCC = "test@acme.com"
objMessage.TextBody = "This is some sample message text." 
objMessage.Send
