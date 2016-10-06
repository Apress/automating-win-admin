Const UF_DONTEXPIRE_PASSWD = 65536
Set objUser = GetObject("LDAP://CN=Fred Smith,CN=Users,DC=Acme,DC=com")
 objUser.userAccountControl = objUser.userAccountControl Or UF_DONTEXPIRE_PASSWD
objUser.SetInfo
