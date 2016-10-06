strDNSServer = "dnsserver.acme.com"
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")

Set objNames = objService.ExecQuery("Select * FROM MicrosoftDNS_CNAMEType Where OwnerName='www.acme.com'")

For Each objName In objNames
    Wscript.Echo "Modifying address" & objName.OwnerName & " " & objName.DomainName
    objName.Modify 600, "estore.acme.com"
Next
