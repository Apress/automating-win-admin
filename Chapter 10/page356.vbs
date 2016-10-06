strDNSServer = "dnsserver.acme.com"
strDomain = "acme.com"
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")

Set objRecords = objService.ExecQuery("Select * FROM MicrosoftDNS_AType Where DomainName='" & strDomain & "'")

'loop through and delete all A type records in domain Acme.com
For Each objRecord In objRecords
    Wscript.Echo "Deleting record " & objRecord.OwnerName
    objRecord.Delete_
Next

'delete the zone Acme.com on server dnsserver.acme.com
Set objZone = objService.Get("MicrosoftDNS_Zone.ContainerName=""" & strDomain & _
                """,DnsServerName=""" & strDNSServer & """,Name=""" & strDomain & """")
objZone.Delete_
