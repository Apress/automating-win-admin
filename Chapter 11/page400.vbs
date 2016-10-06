strTarget = "www.ibm.com"

Set objPings = GetObject("winmgmts:{impersonationLevel=impersonate}" & _
        "root/cimv2").ExecQuery("SELECT * FROM Win32_PingStatus " & _
        "WHERE Address = '" & strTarget & "' AND Timeout=4000 AND TimeToLive =90 And Buffersize=64 ")

For Each objPing In objPings
    If objPing.StatusCode = 0 Then
        Wscript.Echo strTarget & " is alive "
    Else
        Wscript.Echo strTarget & " is not responding"
        Wscript.Echo "Status code is " & objPing.StatusCode
    End If
Next
