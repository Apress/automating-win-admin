Dim objSession, objShell, strProf

Set objShell = CreateObject("WScript.Shell")
' create a MAPI session 
Set objSession = CreateObject("MAPI.Session")

If InStr(objSession.OperatingSystem, "NT") > 0 Then
  strProf = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\" & _
         "Windows Messaging Subsystem\Profiles\DefaultProfile"
Else
  strProf = "HKCU\Software\Microsoft\Windows Messaging Subsystem" & _
        "\Profiles\DefaultProfile"
End If
'logon using the default profile name
objSession.Logon objShell.RegRead(strProf)
