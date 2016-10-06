'create basic non-AD primary zone for domain on a Windows 2003 server
'connect to DNS server and create DNS zone object

strNewZoneName = "acmeY.com"

'Set objService = GetObject("winmgmts:\\dnsserver.acme.com\root\MicrosoftDNS")
Set objService = GetObject("winmgmts:\\w2k3-01\root\MicrosoftDNS")
Set objZone = objService.Get("MicrosoftDNS_Zone")

objZone.CreateZone strNewZoneName, 0, True, , , "admin@adomain.com"
