Dim objIP
  
 'create IPnetwork object
 Set objIP = CreateObject("SScripting.IPNetwork")
 
'check if machine 'elvis' is 
 If objIP.Ping("elvis") = 0 Then
    WScript.Echo "Elvis is alive!"
 End If
