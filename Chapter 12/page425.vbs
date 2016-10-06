Const cdoSendUsingPort = 2
Const schema = "http://schemas.microsoft.com/cdo/configuration/" 
strSmartHost = "exchange.acme.com"

Set objMsg = CreateObject("CDO.Message")
Set objConf = CreateObject("CDO.Configuration")

Set objFlds = objConf.Fields
'set the CDOSYS configuration fields to use port 25 on the SMTP server
'and use 
With objFlds
  .Item(schema & "sendusing") = cdoSendUsingPort
  .Item(schema & "smtpserver") = strSmartHost
  .Item(schema & "smtpserverport") = 25
  .Update
End With

' apply the settings to the message
With objMsg
Set .Configuration = objConf
  .To = "administrator@acme.com"
  .From = "freds@acme.com"
  .Subject = "Set message subjects"
  .TextBody = "Set body of text"
  .Send
End With
