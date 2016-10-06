'get contact Joe Smith
Set objUser = GetObject("LDAP://cn=Joe SmithCX,ou=Contacts,dc=acme,dc=com ")
'change the target address
objUser.TargetAddress = "joesmith@hotmail.com"
objUser.SetInfo
