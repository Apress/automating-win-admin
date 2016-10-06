strTarget = "216.239.51.104"

Set objPings = GetObject("winmgmts:{impersonationLevel=impersonate}" & _
        "root/cimv2").ExecQuery("SELECT * FROM Win32_PingStatus " & _
        "WHERE Address = '" & strTarget & "' ")

For Each objPing In objPings
    If objPing.StatusCode = 0 Then
        Wscript.Echo strTarget & " is alive "
        Wscript.Echo "Response time  = " & objPing.ResponseTime
        Wscript.Echo "TTL  = " & objPing.ResponseTimeToLive
    Else
        Wscript.Echo strTarget & " is not responding"
        Wscript.Echo "Status code is " & objPing.StatusCode
    End If
Next
