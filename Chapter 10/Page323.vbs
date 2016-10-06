Dim objWMI, nTimeZone

'get a reference to instance of Win32_ComputerSystem class for computer Odin
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}" & _
             "!Win32_ComputerSystem='Thor'")

nTimeZone = objWMI.Currenttimezone
Wscript.Echo "THe timezone is " & nTimeZone & " hours of GMT"