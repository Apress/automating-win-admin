'delete a user using the Active Ddirectory provider
Dim objContainer
'get a reference to the LDAP container that contains the object to delete
Set objContainer = GetObject("LDAP://OU=Accountants,DC=Acme,DC=com")

'delete Fred Smith by specifying his relative distinguished name
objContainer.delete "user", "CN=Fred Smith"
