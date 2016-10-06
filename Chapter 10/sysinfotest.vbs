Dim objSysInfo

Set objSysInfo = CreateObject("ENTWSH.SysInfo")
WScript.Echo "BIOS Version:" & objSysInfo.BIOSVersion
 WScript.Echo "CPU:" & objSysInfo.CPU
 WScript.Echo "Memory:" & objSysInfo.Memory
 WScript.Echo "O/S Version:" & objSysInfo.OS
 WScript.Echo "O/S Registered User:" & objSysInfo.RegisteredUser
 WScript.Echo "O/S Serial #:" & objSysInfo.SerialNumber
 WScript.Echo "Virtual Memory:" & objSysInfo.VirtualMemory
