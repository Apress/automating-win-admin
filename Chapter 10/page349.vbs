'create a primary and secondary zone
strNewZoneName = "acme.com"
strPrimaryServer = "dnsserver.acme.com"
strSecondaryServer = "dnsserver2.acme.com"
strSecondaryIP = "192.168.0.226"

'get DNS service for primary service
Set objService = GetObject("winmgmts:\\" & strPrimaryServer & "\root\MicrosoftDNS")

'create a new zone
Set objZone = objService.Get("MicrosoftDNS_Zone")
objZone.CreateZone strNewZoneName, 0

'get a reference to the newly create zone
Set objZone = objService.Get("MicrosoftDNS_Zone.ContainerName=""" & strNewZoneName & _
                """,DnsServerName=""" & strPrimaryServer & """,Name=""" & strNewZoneName & """")

'set secondary server address
objZone.ResetSecondaries Array(strSecondaryIP), 2

'get WMI service for secondary DNS server and create DNS object
Set objService = GetObject("winmgmts:\\" & strSecondaryServer & "\root\MicrosoftDNS")
Set objZone = objService.Get("MicrosoftDNS_Zone")

aMasterIP = Array(strSecondaryIP)
'create secondary zone pointing to the primary zone
objZone.CreateZone strNewZoneName, 1, strNewZoneName & ".dns", aMasterIP
