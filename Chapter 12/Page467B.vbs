'get all messages that contain sales data and extract custom sales fields
'create a MAPI session,then log on
Set objSession = CreateObject("MAPI.Session")
' supply a valid profile
objSession.Logon "Valid Profile"
'get a reference to the InBox messages
Set objMsgColl = objSession.InBox.Messages
'get a reference to the filter for the InBox messages
Set objFilter = objMsgColl.Filter
'set the filter properties to filter all messages that are unread, of type 
'IPM.Note.SCBWeeklySales 
objFilter.Type = " IPM.Note.SCBWeeklySales" 
objFilter.Unread = True '  filter only messages that haven't been read
objFilter.Subject = "Weekly Sales" 
 
'loop through the Messages collection and display the SCBWeeklySales 
'and SCBWeeklyWages for any 'messages that meet the filter criteria
Set objMessage = objMsgColl.GetFirst()

Do While Not objMessage Is Nothing
'display the subject for the current message
 WScript.Echo objMessage.Fields("SCBWeeklySales")
 WScript.Echo objMessage.Fields("SCBWeeklyWages")
 'get the next message
 Set objMessage = objMsgColl.GetNext()
Loop
objSession.Logoff
