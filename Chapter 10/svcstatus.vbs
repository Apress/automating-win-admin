'svcstatus.vbs
'Description:monitors changes in service status
 Option Explicit
 Dim objService, objEvent, objEvents

 Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}")

 'query for any changes in services every 60 seconds
 Set objEvents = objService.ExecNotificationQuery _ 
      ("select * from __instancemodificationevent within 60 " & _ 
        "where targetinstance isa 'Win32_Service'") 

 WScript.Echo "Waiting for service change..."
 Do 
  Set objEvent = objEvents.Nextevent
  WScript.Echo objEvent.TargetInstance.Description & " state changed to " & _
                objEvent.TargetInstance.State & " at " & Now
 Loop
