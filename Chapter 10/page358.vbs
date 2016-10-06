strDNSServer = "dnsserver.acme.com"
strNewSubnet = "172.30.1."
strPTRContainer = "1.30.172.in-addr.arpa"

Set objService = GetObject("winmgmts:\\" & strDNSServer & "\root\MicrosoftDNS")

Set objPTRItem = objService.Get("MicrosoftDNS_PTRType")
Set objARecords = objService.ExecQuery("Select * FROM MicrosoftDNS_AType Where DomainName='acme.com'")

For Each objARecord In objARecords

    If Left(objARecord.OwnerName, 4) = "wkst" Then
        Wscript.Echo objARecord.OwnerName & " " & objARecord.IPAddress & " " & objARecord.DomainName
        'get the last octet from the current record
        strOctet = Mid(objARecord.IPAddress, 11)
        strNewIP = strNewSubnet & strOctet
        
        'modify the IP address to 172 subnet
        objARecord.Modify 600, strNewIP
    
        'create corresponding PTR record
        objPTRItem.CreateInstanceFromPropertyData _
            strDNSServer, strPTRContainer, strOctet & "." & strPTRContainer, _
            intRecordClass, nTTL, objARecord.OwnerName
    
    End If
Next
