Set objUser = GetObject("LDAP://cn=Fred,cn=Users,dc=c3i,dc=com")
 objUser.DeleteMailbox
 objUser.SetInfo
