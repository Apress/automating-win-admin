<?xml version="1.0" ?>
<job>
<!--comment
Script:postpublic.wsf
-->
 <script language="VBScript" src="maillib.vbs">
 <![CDATA[
 ' create a MAPI session,then log on
  Set objSession = CreateObject("MAPI.Session")

  ' supply a valid profile
  objSession.Logon "SCB"

  'get a Reference to the Public Folders folder Public Posting 
  Set objFolder = GetFolderObj(objSession, _
        "Public Folders\All Public Folders\Public Posting") 

  'add a new message to the Public Folder. Note that no recipient is 'required.
  Set objMessage = objFolder.Messages.Add("New Post", "Testing Posting")

  objMessage.TimeReceived = Now 
  objMessage.TimeSent = Now 
  objMessage.Unread = True  
  objMessage.Sent = True  
  objMessage.Submitted = False   
  'set type as a posted message - not required, but message will appear 
  'with a small 'posted' icon
  objMessage.Type = "IPM.Post" 
  objMessage.Update 
  objSession.Logoff
  ]]>
  </script>
</job>

