strComputer = "Loki" 
Set objService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
'get computer object and determine timezone
Set objComputer = objService.Get("Win32_ComputerSystem='" & strComputer & "'")
nTimeZone = objComputer.Currenttimezone

If nTimeZone >= 0 Then 
  strTimeZone = "+" & nTimeZone
Else
  strTimeZone = "-" & nTimeZone
End If

'get local time on remote computer
Set objTime = objService.Get("Win32_LocalTime=@")

'add 30 seconds to the current time on remote computer
dTm = TimeSerial (objTime.Hour,objTime.Minute,objTime.Second + 30)

strTime = Right("0" & Hour(dTm),2) & _
          Right("0" & Minute(dTm),2) & _
          Right("0" & Second(dTm),2) 

nDay = 2 ^ (Day(Date) -1)

Set objNewJob = objService.Get("Win32_ScheduledJob")
'create a one off scheduled job that will run 30 seconds in future
 errJobCreated = objNewJob.Create _
    ("notepad.exe", "********" & strTime  & ".000000" & strTimeZone, _
    False, ,nDay True, JobID)

