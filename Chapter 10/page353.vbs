strDNSServer = "dnsserver.acme.com"
strContainer = "acme.com"
strOwnerAddress = "mail.acme.com"
intRecordClass = 1
nTTL = 600

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objItem = objService.Get("MicrosoftDNS_MXType")
 
'create two MX records
objItem.CreateInstanceFromPropertyData _
    strDNSServer, strContainer, strOwnerAddress, intRecordClass, nTTL, 10, "mail1.acme.com"

objItem.CreateInstanceFromPropertyData _
    strDNSServer, strContainer, strOwnerAddress, intRecordClass, nTTL, 20, "mail2.acme.com"
