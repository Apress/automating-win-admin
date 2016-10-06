Set objUser = GetObject("LDAP://CN=fred smith,CN=Users,DC=Acme,DC=com")
objUser. msRADIUSFramedIPAddress= 1 * 16777216 + 2 * 65536 + 3  * 256 + 4
objUser.SetInfo
