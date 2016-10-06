strDNSServer = "dnsserver.acme.com"
strContainer = "acme.com"
strAddress = "www.acme.com"
intRecordClass = 1
nTTL = 600
strPrimaryName = "eshop.acme.com"

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objItem = objService.Get("MicrosoftDNS_CNAMEType")
 
objItem.CreateInstanceFromPropertyData _
    strDNSServer, strContainer, strAddress, intRecordClass, nTTL, strPrimaryName
