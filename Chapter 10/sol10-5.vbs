Const WMICConst = "winmgmts:{impersonationLevel=impersonate}!"

Dim objWMINetConfig, nResult

'get the instance for the first adapater
Set objWMINetConfig = _ 
    GetObject(WMICConst & "Win32_NetworkAdapterConfiguration.Index=1")

'assign two static IP addresses to the adapter 
nResult=objWMINetConfig.EnableStatic(Array("192.168.1.2","192.168.1.3"), _ 
                             Array("255.255.255.0", "255.255.255.0"))
