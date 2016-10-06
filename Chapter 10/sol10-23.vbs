strDNSServer = "dnsserver.acme.com"
strContainer = "acme.com"
strOwnerAddress = "shop1.acme.com"
intRecordClass = 1
nTTL = 600
strIPAddress = "192.168.1.101"

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objAType = objService.Get("MicrosoftDNS_AType")
objAType.CreateInstanceFromPropertyData  strDNSServer, strContainer, strOwnerAddress, intRecordClass, nTTL, strIPAddress
