'extract.vbs
Dim objSession, objMsgColl, objMessage, strFile
Set objSession = CreateObject("MAPI.Session")
' supply a valid profile
objSession.Logon "Valid Profile"

'get a reference to the InBox messages
Set objMsgColl = objSession.InBox.Messages 
'filter all unread messages
objMsgColl.Filter.Unread = True 

'loop through the Messages collection and extract attachment for any messages 
'that meet the filter criteria
Set objMessage = objMsgColl.GetFirst()

Do While Not objMessage Is Nothing
 WScript.Echo objMessage.Subject, objMessage.Attachments.Count
   For Each objAttachment In objMessage.Attachments
    strFile = objAttachment.Source
      If strFile <> "" Then 
      strFile = Mid(strFile,InstrRev(strFile,"\") + 1)
      objAttachment.WriteToFile "d:\data\" & strFile
     End If  
   Next
   objMessage.Unread = False ' flag message as read
   objMessage.Update 'update message
  'get the next message
  Set objMessage = objMsgColl.GetNext()
Loop
objSession.Logoff
