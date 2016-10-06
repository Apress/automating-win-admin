'create basic non-ADSI reverse lookup zone on a Windows 2003 server 
strDNSServer = "dnsserver.acme.com"

'connect to DNS server and create DNS reverse zone 
Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objZone = objService.Get("MicrosoftDNS_Zone")

objZone.CreateZone "1.168.192.in-addr.arpa", 0
