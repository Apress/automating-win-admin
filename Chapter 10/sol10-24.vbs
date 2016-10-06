strDNSServer = "w2k3-01" '"dnsserver.acme.com"
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")

Set objNames = objService.ExecQuery("Select * FROM MicrosoftDNS_AType Where OwnerName='wkst1.acme.com'")

For Each objName In objNames
    objName.Delete_
Next
