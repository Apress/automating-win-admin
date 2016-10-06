Set objDir = GetObject("IIS://thor/W3SVC/1/ROOT/images")
'if access is granted by default list the denied IP addresses and domains,
'otherwise list the granted IP adresses and domains
If objDir.IPSecurity.GrantByDefault Then
    Wscript.Echo "IP Addresses denied access"
    For Each obj In objDir.IPSecurity.IPDeny
        Wscript.Echo obj
    Next
    Wscript.Echo "Domains denied access"
    For Each obj In objDir.IPSecurity.DomainDeny
        Wscript.Echo obj
    Next
Else
    Wscript.Echo "IP Addresses granted access"
    For Each obj In objDir.IPSecurity.IPGrant
        Wscript.Echo obj
    Next
    Wscript.Echo "Domains granted access"
    For Each obj In objDir.IPSecurity.DomainGrant
        Wscript.Echo obj
    Next
End If
