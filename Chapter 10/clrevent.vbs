'clrevent.vbs
'backs up and clears local event log if number of records is 
'greater than 5000 records

Dim objWMISet, objInstance, nResult, strFile


'get an instance of a WMI Service object
Set objService = _
        GetObject("winmgmts:{impersonationLevel=impersonate,(Backup)}")

'get the instances of the Win32_NTEventLogFile class
Set objWMISet = objService.InstancesOf("Win32_NTEventLogFile")

 For Each objInstance In objWMISet

 'check if number of event records is more than 5000, clear log
   If objInstance.NumberOfRecords > 5000 Then

   strFile = "d:\eventlogs\backup" & objInstance.LogFileName _
                                 & " " & MonthName(Month(Date)) & _
                                 "-" & Day(Date) & "-" & Year(Date) & _
                                 " " & Hour(Now) & "-" & _ 
                                 Minute(Now) & ".evt"

   'backup and clear log    
   nResult = objInstance.BackupEventlog(strFile)
   
   'check if operation successful
    If nResult<> 0 Then 
	 Wscript.Echo objInstance.LogFileName & " backed up to " & strFile
    End If

   End If

 Next

 Set objInstance = Nothing
 Set objWMISet = Nothing