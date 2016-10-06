Const vbSingle = 4
objSession.Logon "Valid Profile"
Set objMessage = objSession.Outbox.Messages.Add 'create a new message

objMessage.Subject = "Weekly Sales"
objMessage.Text = "Weekly Sales Message"
objMessage.Type = "IPM.Note.SCBWeeklySales" 'set the message type 
Set objField = objMessage.Fields.Add("SCBWeeklySales", vbSingle,143545.50)
Set objField = objMessage.Fields.Add("SCBWeeklyWages", vbSingle, 2333.50)
