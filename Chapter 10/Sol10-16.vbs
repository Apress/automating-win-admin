'connect to Services object for local computer
Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}")

'execute a event notification query to generate events when a new 
'event record is created
Set objEvents = objServices.ExecNotificationQuery _
                     ("select * from __instancecreationevent " & _ 
                      "where targetinstance isa 'Win32_NTLogEvent'") 

'check if error occurred during execution of
if Err <> 0 then
   WScript.Echo Err.Description, Err.Number, Err.Source
End If 

'wait forever..
WScript.Echo "Waiting for NT Events..."
Do 
   Set NTEvent = objEvents.nextevent 
   'check if error occurred
   If Err <> 0 Then
      WScript.Echo Err.Number, Err.Description, Err.Source
      Exit Do
   Else      
      WScript.Echo NTEvent.TargetInstance.Message
   End if
Loop
