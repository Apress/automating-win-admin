strDNSServer = "dnsserver.acme.com"
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")

Set objNames = objService.ExecQuery("Select * FROM MicrosoftDNS_AType Where DomainName='acme.com'")

For Each objName In objNames
    Wscript.Echo objName.OwnerName & " " & objName.IPAddress
Next
