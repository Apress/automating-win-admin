'get a reference to the Win32Process class object on specified machine
Set objProcess=GetObject("winmgmts:{impersonationLevel=impersonate}!" & _
               "\\Odin\root\cimv2:Win32_Process" )
 
  'create process on remote machine
 nResult = objProcess.Create("notepad.exe", , ,nProcID)

 WScript.Echo "The PID for the new instance of notepad is " & nProcID
