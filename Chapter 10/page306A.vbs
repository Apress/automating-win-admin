Const Friday = 16
Const Wednesday = 4
strComputer = "Thor" 

Set objService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

'get computer object and determine timezone
Set objComputer = objService.Get("Win32_ComputerSystem='" & strComputer & "'")
nTimeZone = objComputer.Currenttimezone

If nTimeZone >= 0 Then 
  strTimeZone = "+" & nTimeZone
Else
  strTimeZone = "-" & nTimeZone
End If

Set objNewJob = objService.Get("Win32_ScheduledJob")
'schedule job to repeatedly run on Wednesday and Friday at 10PM
 errJobCreated = objNewJob.Create _
    ("e:\data\backup.bat", "********220000.000000" & strTimeZone, _
    False, Friday + Wednesday , , True, JobID)
