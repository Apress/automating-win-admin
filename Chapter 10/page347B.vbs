'create basic non-ADSI primary zone for domain on a Windows 2000 server
'connect to DNS server and create DNS zone object
Set objService = GetObject("winmgmts:\\w2k-01\root\MicrosoftDNS")
Set objZone = objService.Get("MicrosoftDNS_Zone")

objZone.CreateZone "adomainW2K.com", 1
