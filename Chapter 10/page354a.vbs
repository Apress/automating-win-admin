
strDNSServer = "dnsserver.acme.com"
strContainer = "168.192.in-addr.arpa"
strIPAddress = "101.1.168.192.in-addr.arpa"

intRecordClass = 1
nTTL = 600

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objItem = objService.Get("MicrosoftDNS_PTRType")
 
errResult = objItem.CreateInstanceFromPropertyData _
    (strDNSServer, strContainer, strIPAddress, intRecordClass, nTTL, "www.acme.com")
	
