Dim objFolder, objSession, objMessage, objMessages
Set objSession = CreateObject("MAPI.Session")
objSession.Logon " Profile Name " 

Set objFolder = objSession.InBox ' get a reference to the InBox Folder object
' return the Messages collection for the InBox
Set objMessages = objFolder.Messages ' return the Messages collection for the InBox
'get the first message in the folder that is of message type IPM.Note .
Set objMessage = objMessages.GetFirst("IPM.Note")
' loop through all messages in the InBox folder, displaying the subject
Do While Not objMessage Is Nothing
    WScript.Echo objMessage.Subject 'display the subject of the message
    'get the next message 
     Set objMessage = objMessages.GetNext
Loop
objSession.Logoff
