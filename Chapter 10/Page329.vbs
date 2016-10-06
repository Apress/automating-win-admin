Dim objEvt

'get a reference to the Application event log. The key for the 
'Win32_NTEventLogFile class is the Name property, which represents the 
'path to the event log.
Set objEvt = _
  GetObject("winmgmts:{impersonationLevel=impersonate,(Backup,Security)}" & _
  "!Win32_NTEventLogFile.Name='C:\WINNT40\System32\config\AppEvent.Evt'")

 objEvt.ClearEventlog ("D:\Data\EvenBackup\AppBackup.evt")
