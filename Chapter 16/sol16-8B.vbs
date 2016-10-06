Dim objAddressContainer, objNewAddressContainer
'get a reference to the container the recipient list will reside in...
Set objAddressContainer = _
GetObject(“LDAP://CN=All Address Lists,” & _
“CN=Address Lists Container,CN=C3I,CN=Microsoft Exchange,” & _
“CN=Services,CN=Configuration,DC=c3i,DC=com”)

'create an object of addressBookContainer class, specifying the name
'using in RDN format
Set objNewAddressContainer = _
objAddressContainer.Create(“addressBookContainer”, “CN=Company X”)

'set display name
objNewAddressContainer.DisplayName = “ Company X List”
objNewAddressContainer.instanceType = 4

'set the LDAP search criteria for the address list. The following criteria
' lists all contacts that are employees of Company XYZ
objNewAddressContainer.purportedSearch = “(&(&(&(& (mailnickname=*)(|” & _
“(&(objectCategory=person)(objectClass=contact)))))” & _
“(objectCategory=user)(company=Company X)))”
objNewAddressContainer.systemFlags = 1610612736
objNewAddressContainer.showInAdvancedViewOnly = True
objNewAddressContainer.SetInfo
