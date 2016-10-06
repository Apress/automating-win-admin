
'create 100 addresses in the domain Acme.com
strDNSServer = "dnsserver.acme.com"
strContainer = "acme.com"
strPTRContainer = "1.168.192.in-addr.arpa"
intRecordClass = 1
nTTL = 600

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")
Set objItem = objService.Get("MicrosoftDNS_AType")
Set objPTRItem = objService.Get("MicrosoftDNS_PTRType")
 

For nF = 1 To 100

    strOwner = "wkst" & nF & ".acme.com"
    errResult = objItem.CreateInstanceFromPropertyData (strDNSServer, strContainer, strOwner, intRecordClass, nTTL, "192.168.1." & nF)

    'create corresponding PTR record
    errResult = objPTRItem.CreateInstanceFromPropertyData _
    (strDNSServer, "168.192.in-addr.arpa", nF & "." & strPTRContainer, _
    intRecordClass, nTTL, strOwner)

Next
