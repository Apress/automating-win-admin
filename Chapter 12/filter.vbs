'filter.vbs
Set objSession = CreateObject("MAPI.Session")

objSession.Logon "SCB"
Set objMsgColl = objSession.InBox.Messages
Set objFilter = objMsgColl.Filter 

'set the filter properties to filter all messages that are unread, of 
'type IPM.Note and were received over 6 months a go
objFilter.Type = "IPM.Note" 
objFilter.Unread = True 
objFilter.TimeLast = DateAdd("d", -180, Date) 

' loop through the Messages collection and display the subject 
' for any messages that meet the filter criteria
Set objMessage = objMsgColl.GetFirst()
Do While Not objMessage Is Nothing
 'display the subject for the current message..
 WScript.Echo objMessage.Subject
'get the next message
 Set objMessage = objMsgColl.GetNext()
Loop
objSession.Logoff
