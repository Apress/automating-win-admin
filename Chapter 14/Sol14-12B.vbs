'get a reference to the Auditors group under the Accounting organizational unit
Set objGroup = GetObject("LDAP://cn=Auditors,ou=Accounting,DC=Acme,DC=com")

'display name of each object in group
For Each objUser In objGroup.Members
  Wscript.Echo objUser.Name 
Next
