Dim objServices, objWMIObject, nResult
Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}")

Set objWMIObject = objServices.Get("Win32_Process.Handle=326")
nResult = objProcess.Terminate(0)
