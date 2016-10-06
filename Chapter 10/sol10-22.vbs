Set objZone = GetObject("winmgmts:\\dnsserver.acme.com\root\MicrosoftDNS:MicrosoftDNS_Zone.ContainerName=""acme.com"",DnsServerName=""dnsserver.acme.com"",Name=""acme.com""")

objZone.DataFile = "acemcomzone.dns"
objZone.NoRefreshInterval = 34
objZone.RefreshInterval = 75
objZone.Put_
