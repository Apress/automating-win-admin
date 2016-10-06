'create basic non-ADSI primary zone on a Windows 2003 server
strNewZoneName = "acme.com"
strDNSServer = "dnsserver.acme.com"

'connect to DNS server and create DNS zone object
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objZone = objService.Get("MicrosoftDNS_Zone")

objZone.CreateZone strNewZoneName, 0
