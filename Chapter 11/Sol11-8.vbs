Dim objIP, strAddress
'create an instance of the scripting host object
Set objIP = CreateObject("SScripting.IPNetwork")
'lookup the domain name for a IP address
strAddress = objIP.DNSLookup("207.46.230.219")

WScript.Echo "The FQDN for the address is " & strAddress
