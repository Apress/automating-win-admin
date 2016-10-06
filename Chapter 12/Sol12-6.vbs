Const MsgHigh = 2

Dim objSession, objMessage 
Set objSession = CreateObject("MAPI.Session")
objSession.Logon 
Set objMessage = objSession.Outbox.Messages.Add

objMessage.Subject = "Test message" 
objMessage.Importance = MsgHigh
objMessage.Sensitivity =  True 
objMessage.ReadReceipt = True 
objMessage.DeliveryReceipt = True objMessage.Type = "IPM.SCBSpecialMessage" 
objMessage.Text= "Testing" 
